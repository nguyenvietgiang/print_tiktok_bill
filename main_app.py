"""
TikTokPrint — Desktop App TikTok Seller Automation + Bill Calculate
====================================================================
GUI: PySide6 (Qt for Python) — 4-tab layout with QSS stylesheet (SaaS Light Theme)
Chạy: python main_app.py
Đóng gói .exe: pyinstaller TTS_Bill.spec
"""
import os, sys, json, shutil, threading, re as _re
from datetime import datetime, timedelta
from pathlib import Path

# ═══════════════════════════════════════════════════════════
# PySide6 imports
# ═══════════════════════════════════════════════════════════
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QGroupBox, QCheckBox, QRadioButton,
    QButtonGroup, QSpinBox, QComboBox, QTextEdit, QListWidget,
    QListWidgetItem, QScrollArea, QStackedWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QFileDialog, QMessageBox,
    QFrame, QApplication,
)
from PySide6.QtCore import (
    Qt, Signal, Slot, QThread, QTimer, QObject, QMetaObject,
)
from PySide6.QtGui import (
    QFont, QTextCursor,
)

from playwright.sync_api import sync_playwright

# ═══════════════════════════════════════════════════════════
# Paths & config
# ═══════════════════════════════════════════════════════════
BASE_DIR = Path(__file__).parent
BILL_DIR = BASE_DIR / 'bill_calculate'
UPLOAD_DIR = BILL_DIR / 'uploads'
sys.path.insert(0, str(BILL_DIR))

try:
    from calculator import process_all
except ImportError:
    process_all = None

# ═══════════════════════════════════════════════════════════
# Frozen / source mode — detect paths
# ═══════════════════════════════════════════════════════════
if getattr(sys, 'frozen', False):
    _portable_browsers = os.path.join(os.path.dirname(sys.executable), 'ms-playwright')
    if os.path.isdir(_portable_browsers):
        os.environ['PLAYWRIGHT_BROWSERS_PATH'] = _portable_browsers
    else:
        os.environ.setdefault('PLAYWRIGHT_BROWSERS_PATH',
            os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'ms-playwright'))
    BASE_DIR = Path(sys.executable).parent
    BILL_DIR = BASE_DIR / 'bill_calculate'
    UPLOAD_DIR = BILL_DIR / 'uploads'
    DEFAULT_COOKIE = Path(sys._MEIPASS) / 'seller-vn.tiktok.com_25-06-2026.json'
    MASTER_DEFAULT = Path(sys._MEIPASS) / 'mã combo.xlsx'
    RETAIL_DEFAULT = Path(sys._MEIPASS) / 'sp bán lẻ.xlsx'
else:
    os.environ.setdefault('PLAYWRIGHT_BROWSERS_PATH',
        os.path.join(os.path.expanduser('~'), 'AppData', 'Local', 'ms-playwright'))
    DEFAULT_COOKIE = BASE_DIR / 'seller-vn.tiktok.com_25-06-2026.json'
    MASTER_DEFAULT = BASE_DIR / 'mã combo.xlsx'
    RETAIL_DEFAULT = BASE_DIR / 'sp bán lẻ.xlsx'

TARGET_URL = 'https://seller-vn.tiktok.com'
ORDERS_URL = 'https://seller-vn.tiktok.com/order?order_status%5B%5D=1&selected_sort=11&tab=to_ship&page_size=50'
CARRIER_URLS = {
    'J&T':           ORDERS_URL + '&shipping_provider_id%5B%5D=6841743441349706241',
    'GHN':           ORDERS_URL + '&shipping_provider_id%5B%5D=7252807945006614278',
    'VietNam Post':  ORDERS_URL + '&shipping_provider_id%5B%5D=7062208235196909313',
    'Best Express':  ORDERS_URL + '&shipping_provider_id%5B%5D=7099655686241388293',
    'Viettel Post':  ORDERS_URL + '&shipping_provider_id%5B%5D=7155825439565416197',
    'J&T Cargo VN':  ORDERS_URL + '&shipping_provider_id%5B%5D=7581675938962736917',
}
BATCH_SIZE = 50
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ============================================================
# AUTOMATION
# ============================================================
def run_automation(cookie_path, output_dir, max_orders, log_cb, state_cb, stop_event=None,
                   existing_playwright=None, existing_browser=None, carrier=None, test_mode=False):
    with open(cookie_path, 'r', encoding='utf-8') as f:
        cd = json.load(f)
    cookies_list = cd.get('cookies', cd if isinstance(cd, list) else [])
    pdf_files = []
    total_printed = 0
    target = max_orders if max_orders > 0 else 10**9
    use_select_all = (max_orders == 0)
    batch_num = 0
    carrier_label = f' [{carrier}]' if carrier else ''
    orders_url = CARRIER_URLS.get(carrier, ORDERS_URL)
    if stop_event is None: stop_event = threading.Event()

    # ── Khởi tạo / tái sử dụng browser ──
    browser_ok = False
    if existing_browser and existing_playwright:
        try:
            playwright = existing_playwright
            browser = existing_browser
            if not browser.is_connected():
                log_cb('⚠ Browser cũ đã ngắt kết nối — tạo mới...', 'warn')
                raise RuntimeError('browser disconnected')
            # Dùng context có sẵn (giữ cookie session)
            context = browser.contexts[0] if browser.contexts else browser.new_context(
                viewport={'width': 1366, 'height': 768},
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36',
                accept_downloads=True)
            # Tạo page MỚI TRƯỚC (để context không bao giờ rỗng), rồi mới đóng page cũ
            page = context.new_page()
            log_cb('♻ Dùng lại browser — tạo tab mới...', 'info')
            # Đóng page cũ SAU (dùng list() để tránh lỗi modify khi duyệt)
            for p in list(context.pages):
                if p != page:
                    try:
                        if not p.is_closed():
                            p.close()
                    except Exception:
                        pass
            page.goto(orders_url, wait_until='networkidle', timeout=60000)
            page.wait_for_timeout(4000)
            browser_ok = True
        except Exception:
            log_cb('⚠ Không dùng lại được browser cũ — tạo mới...', 'warn')
            try:
                existing_browser.close()
            except: pass
            try:
                existing_playwright.stop()
            except: pass
            existing_browser = None
            existing_playwright = None

    if not browser_ok:
        playwright = sync_playwright().start()
        browser = playwright.chromium.launch(headless=False, args=['--disable-blink-features=AutomationControlled'])
        context = browser.new_context(viewport={'width': 1366, 'height': 768},
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36',
            accept_downloads=True)
        pw_cookies = []
        for c in cookies_list:
            if not c.get('name') or not c.get('value'): continue
            pw = {'name': c['name'], 'value': c['value'], 'domain': c.get('domain', '.tiktok.com'), 'path': c.get('path', '/')}
            if c.get('expirationDate') and not c.get('session'): pw['expires'] = int(c['expirationDate'])
            if 'httpOnly' in c: pw['httpOnly'] = c['httpOnly']
            if 'secure' in c: pw['secure'] = c['secure']
            if c.get('sameSite'):
                pw['sameSite'] = {'strict':'Strict','lax':'Lax','no_restriction':'None','unspecified':'Lax'}.get(c['sameSite'],'Lax')
            pw_cookies.append(pw)
        context.add_cookies(pw_cookies)
        page = context.new_page()
        page.goto(TARGET_URL, wait_until='domcontentloaded', timeout=30000)
        page.wait_for_timeout(2000)

    try:
        while total_printed < target:
            if stop_event.is_set():
                log_cb('⏹ Đã dừng theo yêu cầu.', 'warn'); break
            batch_num += 1
            batch_target = min(BATCH_SIZE, target - total_printed)
            log_cb(f'▸ Batch {batch_num}{carrier_label}: chọn {batch_target} đơn (đã in {total_printed}/{target})', 'batch')

            state_cb('navigating', f'Batch {batch_num}: Đang tải danh sách đơn...')
            # Retry 3 lần khi mất mạng hoặc timeout
            MAX_GOTO_RETRIES = 3
            for goto_attempt in range(MAX_GOTO_RETRIES):
                try:
                    page.goto(orders_url, wait_until='networkidle', timeout=60000)
                    break
                except Exception as e:
                    if goto_attempt < MAX_GOTO_RETRIES - 1:
                        log_cb(f'  ⚠ Lỗi mạng (lần {goto_attempt+1}/{MAX_GOTO_RETRIES}): {e} — thử lại sau 5s...', 'warn')
                        page.wait_for_timeout(5000)
                    else:
                        log_cb(f'  ✗ Thất bại sau {MAX_GOTO_RETRIES} lần thử: {e}', 'err')
                        raise
            page.wait_for_timeout(4000)

            total_avail = page.evaluate("() => document.querySelectorAll('td.col-checkbox label.p-checkbox').length")
            if total_avail == 0:
                # Kiểm tra có phải do chưa đăng nhập hay thực sự hết đơn
                try:
                    is_logged_out = page.evaluate("""() => {
                        return (document.querySelector('input[name="email"]') !== null ||
                                document.querySelector('input[type="email"]') !== null ||
                                window.location.href.includes('login') ||
                                window.location.href.includes('signin') ||
                                document.querySelector('[data-testid="login"]') !== null);
                    }""")
                except Exception:
                    is_logged_out = False
                if is_logged_out:
                    log_cb('  ✗ Cookie hết hạn hoặc chưa đăng nhập — vui lòng cập nhật cookie.', 'err')
                else:
                    log_cb('  ✓ Hết đơn khả dụng — hoàn thành.', 'ok')
                break

            if use_select_all:
                log_cb(f'  🖱 Đang chọn tất cả đơn hàng...', 'info')
                header_clicked = False
                for hdr_sel in [
                    'th svg.p-checkbox-mask-icon', 'th .p-checkbox-mask-icon', 'th .p-checkbox',
                    'th.col-checkbox .p-checkbox', 'th.col-checkbox label.p-checkbox', 'th input[type="checkbox"]',
                ]:
                    try:
                        hdr = page.locator(hdr_sel).first
                        if hdr.count() > 0 and hdr.is_visible(timeout=2000):
                            hdr.click(); header_clicked = True
                            log_cb(f'  ✓ Đã click checkbox header ({hdr_sel})', 'ok'); break
                    except: pass
                if not header_clicked:
                    header_clicked = page.evaluate('''() => {
                        const svg = document.querySelector('th svg.p-checkbox-mask-icon, th .p-checkbox-mask-icon svg, th svg[class*="checkbox"]');
                        if (svg) {
                            const parent = svg.closest('.p-checkbox') || svg.closest('label') || svg.closest('th');
                            if (parent) { parent.click(); return true; }
                            svg.dispatchEvent(new MouseEvent('click', {bubbles: true})); return true;
                        }
                        const cb = document.querySelector('th .p-checkbox, th.col-checkbox label, thead input[type="checkbox"]');
                        if (cb) { cb.click(); return true; }
                        return false;
                    }''')
                    if header_clicked: log_cb('  ✓ Đã click checkbox header (JS fallback)', 'ok')
                page.wait_for_timeout(2500)

                select_all_btn = None
                for sel_text in ['Chọn tất cả', 'Select all', 'Chọn tất', 'Select All']:
                    for tag in ['span', 'button', 'div']:
                        try:
                            btns = page.locator(f'{tag}:has-text("{sel_text}")')
                            cnt = btns.count()
                            for j in range(cnt):
                                b = btns.nth(j)
                                if b.is_visible(timeout=800): select_all_btn = b; break
                        except: pass
                        if select_all_btn: break
                    if select_all_btn: break

                if select_all_btn:
                    select_all_btn.click(timeout=5000)
                    log_cb(f'  ✅ Đã bấm "Chọn tất cả"', 'ok')
                    page.wait_for_timeout(2000)
                    checked = page.evaluate("() => document.querySelectorAll('input[type=\"checkbox\"]:checked').length")
                    log_cb(f'  ✓ Đã chọn {checked} đơn hàng', 'ok')
                else:
                    # Không có nút "Chọn tất cả" → có thể header checkbox đã chọn hết rồi (≤50 đơn)
                    already_checked = page.evaluate("() => document.querySelectorAll('input[type=\"checkbox\"]:checked').length")
                    if already_checked > 0:
                        checked = already_checked
                        log_cb(f'  ✓ Header checkbox đã chọn {checked} đơn (không cần nút "Chọn tất cả")', 'ok')
                    else:
                        log_cb('  ⚠ Không tìm thấy nút "Chọn tất cả" — fallback tick tay', 'warn')
                        to_select = min(batch_target, total_avail)
                        cbs = page.query_selector_all('td.col-checkbox label.p-checkbox')
                        checked = 0
                        for cb in cbs[:to_select]:
                            try: cb.click(); checked += 1; page.wait_for_timeout(120)
                            except: pass
                        page.wait_for_timeout(800)
                        log_cb(f'  ✓ Đã tick {checked}/{to_select} đơn (fallback)', 'ok')
                force_stop = True
                batch_target = 10**9
            else:
                to_select = min(batch_target, total_avail)
                cbs = page.query_selector_all('td.col-checkbox label.p-checkbox')
                checked = 0
                for cb in cbs[:to_select]:
                    try: cb.click(); checked += 1; page.wait_for_timeout(120)
                    except: pass
                page.wait_for_timeout(800)
                force_stop = total_avail < batch_target
                log_cb(f'  ✓ Đã chọn {checked}/{batch_target} đơn', 'ok')

            if checked == 0:
                log_cb('  ✓ Hết đơn — hoàn thành.', 'ok'); break

            if test_mode:
                log_cb('  🧪 TEST MODE: Dừng tại bước chọn đơn — không in.', 'warn')
                try:
                    ss = str(Path(output_dir) / f'test_mode_batch{batch_num}.png')
                    page.screenshot(path=ss); log_cb(f'  📸 Screenshot: {ss}', 'info')
                except: pass
                total_printed += checked; log_cb(f'  📊 Đã chọn {checked} đơn (test mode — không in)', 'info'); break

            state_cb('printing', f'Batch {batch_num}: Đang in...')
            ship_btn = None
            for btn_text in ['Sắp xếp vận chuyển và in', 'Arrange shipment and print', 'Sắp xếp vận chuyển']:
                try:
                    btn = page.locator(f'button:has-text("{btn_text}")').first
                    if btn.count() > 0 and btn.is_visible(timeout=3000): ship_btn = btn; break
                except: pass
            if not ship_btn:
                all_btns = page.locator('button').all()
                for b in all_btns:
                    try:
                        txt = b.inner_text().strip().lower()
                        if 'vận chuyển' in txt and 'in' in txt: ship_btn = b; break
                    except: pass
            if not ship_btn:
                log_cb('  ✗ KHÔNG TÌM THẤY nút "Sắp xếp vận chuyển và in"!', 'err')
                try:
                    ss = str(Path(output_dir) / f'debug_no_ship_btn_batch{batch_num}.png')
                    page.screenshot(path=ss); log_cb(f'  📸 Screenshot: {ss}', 'info')
                except: pass
                total_printed += checked; break
            ship_btn.click()
            log_cb('  ✓ Đã bấm "Sắp xếp vận chuyển và in"', 'ok')
            page.wait_for_timeout(3000)

            state_cb('printing', f'Batch {batch_num}: Đợi popup "Tiếp theo"...')
            tieptheo_btn = None
            for _ in range(10):
                for selector in ['button:has-text("Tiếp theo")', 'button:has-text("Next")']:
                    try:
                        btns = page.locator(selector)
                        cnt = btns.count()
                        for j in range(cnt):
                            b = btns.nth(j)
                            if b.is_visible(timeout=500): tieptheo_btn = b; break
                    except: pass
                    if tieptheo_btn: break
                if tieptheo_btn: break
                page.wait_for_timeout(1000)

            if tieptheo_btn:
                tieptheo_btn.click(timeout=5000)
                log_cb('  ✓ Đã bấm "Tiếp theo"', 'ok')
                page.wait_for_timeout(3000)

                state_cb('printing', f'Batch {batch_num}: Chọn loại chứng từ...')
                page.wait_for_timeout(2000)
                for doc_label in ['Danh sách đóng gói', 'Danh sách lấy hàng']:
                    try:
                        lbl = page.locator('label').filter(has_text=doc_label).first
                        if lbl.count() > 0:
                            inp = lbl.locator('input')
                            if inp.count() > 0 and not inp.is_checked():
                                lbl.click(); page.wait_for_timeout(300)
                                log_cb(f'  ✓ Đã tick: {doc_label}', 'ok')
                    except: pass
                page.wait_for_timeout(1000)

                in_btn = None
                for btn_text in ['In nhãn ngay sau khi vận chuyển', 'In nhãn ngay', 'Print label immediately']:
                    try:
                        btn = page.locator(f'button:has-text("{btn_text}")').first
                        if btn.count() > 0 and btn.is_visible(timeout=2000): in_btn = btn; break
                    except: pass
                if not in_btn:
                    for b in page.locator('button').all():
                        try:
                            txt = b.inner_text().strip().lower()
                            if 'in nhãn' in txt or 'in nhan' in txt: in_btn = b; break
                        except: pass
                if in_btn:
                    in_btn.click(timeout=5000)
                    log_cb('  ✓ Đã bấm "In nhãn ngay"', 'ok')
                    page.wait_for_timeout(3000)
                else: log_cb('  ⚠ Không tìm thấy nút "In nhãn ngay"', 'warn')
            else:
                log_cb('  ⏭ Không có popup "Tiếp theo" → đi thẳng bước tải xuống', 'info')

            state_cb('downloading', f'Batch {batch_num}: Đợi popup "Tải xuống tất cả"...')
            taixuong_btn = None
            for _ in range(300):  # 300 × 1s = 5 phút, đợi TikTok server sinh PDF
                btns = page.locator('button:has-text("Tải xuống tất cả"), button:has-text("Download all"), button:has-text("Tải xuống")')
                cnt = btns.count()
                for j in range(cnt):
                    b = btns.nth(j)
                    try:
                        if b.is_visible(timeout=500): taixuong_btn = b; break
                    except: pass
                if taixuong_btn: break
                page.wait_for_timeout(1000)

            if not taixuong_btn:
                log_cb('  ✗ KHÔNG TÌM THẤY nút "Tải xuống tất cả tập tin" — kiểm tra popup!', 'err')
                try:
                    ss = str(Path(output_dir) / f'debug_no_taixuong_batch{batch_num}.png')
                    page.screenshot(path=ss); log_cb(f'  📸 Screenshot: {ss}', 'info')
                except: pass
                total_printed += checked; break

            log_cb('  📥 Đang bấm nút tải xuống...', 'info')
            downloaded_files = []
            def on_download(dl):
                carrier_prefix = carrier.replace(' ', '_').replace('&', 'n') + '_' if carrier else ''
                base_name = dl.suggested_filename or f'PDF_goc_TTS_batch{batch_num}_{len(downloaded_files)}_{datetime.now().strftime("%m-%d_%H-%M-%S")}.pdf'
                suggested = carrier_prefix + base_name if carrier else base_name
                bp = str(Path(output_dir) / suggested)
                dl.save_as(bp)
                downloaded_files.append(bp)
                log_cb(f'  💾 Đã tải: {Path(bp).name}', 'ok')
            
            page.on('download', on_download)
            taixuong_btn.click(timeout=5000)
            log_cb('  ✓ Đã bấm "Tải xuống tất cả"', 'ok')
            # Chờ download hoàn tất: tối đa 5 phút, nhưng nếu có file + 20s ko có thêm → xong
            idle_ticks = 0
            for _ in range(30):  # 30 × 10s = 5 phút tối đa
                page.wait_for_timeout(10000)  # đợi 10 giây
                if downloaded_files:
                    prev_count = len(downloaded_files)
                    page.wait_for_timeout(3000)  # đợi thêm 3s cho các download khác
                    if len(downloaded_files) == prev_count:
                        idle_ticks += 1
                        if idle_ticks >= 2:  # 26 giây không có download mới → xong
                            log_cb(f'  ✅ Download hoàn tất ({len(downloaded_files)} file)', 'ok')
                            break
                    else:
                        idle_ticks = 0  # reset, vẫn còn download mới
            page.remove_listener('download', on_download)
            pdf_files.extend(downloaded_files)
            log_cb(f'  📥 Tổng cộng {len(downloaded_files)} file đã tải', 'info')
            if not downloaded_files:
                log_cb('  ✗ Không bắt được download nào!', 'err')
                total_printed += checked; break

            total_printed += checked
            log_cb(f'  📊 Tiến độ: {total_printed}/{target} đơn, {len(pdf_files)} file PDF', 'info')

            if force_stop:
                log_cb('  ✓ Đã in hết đơn hiện có — hoàn thành.', 'ok'); break
            if total_printed < target and total_avail > 0:
                page.wait_for_timeout(2000)

        return pdf_files, playwright, browser
    except Exception as e:
        log_cb(f'  ✗ Lỗi: {e}', 'err')
        return pdf_files, playwright, browser

# ============================================================
# CALCULATOR
# ============================================================
def run_calculator(pdf_paths, output_dir, master_path, retail_path, log_cb, carrier=''):
    if process_all is None:
        log_cb('✗ Calculator không khả dụng (thiếu module calculator)', 'err'); return []
    if not Path(master_path).exists(): log_cb(f'✗ Không tìm thấy master_data: {master_path}', 'err'); return []
    out_dir = str(output_dir)
    for p in pdf_paths:
        shutil.copy2(p, str(UPLOAD_DIR / Path(p).name))
    try:
        results = process_all(pdf_paths, out_dir, master_path, retail_path, carrier)
        for r in results:
            log_cb(f'  ✓ {r["rows"]} dòng | Qty={r["tong_qty"]} | Sold={r["tong_sold"]} | Promo={r["tong_promo"]}', 'ok')
            for key, fb in r['files'].items():
                src, dst = Path(fb), Path(out_dir) / Path(fb).name
                if src != dst and src.exists(): shutil.copy2(str(src), str(dst)); r['files'][key] = str(dst)
        return results
    except Exception as e: log_cb(f'  ✗ Lỗi: {e}', 'err'); return []

# ═══════════════════════════════════════════════════════════════
# WORKER & PRINTING FUNCTIONS
# ═══════════════════════════════════════════════════════════════
class AutomationWorker(QObject):
    log_message = Signal(str, str)
    state_changed = Signal(str, str)
    job_completed = Signal(object)
    result_file = Signal(str)

    def __init__(self, parent=None):
        super().__init__(parent)
        self._playwright = None
        self._browser = None
        self._stop_event = threading.Event()
        self._running = False

    @Slot(dict)
    def start_job(self, config: dict):
        self._running = True
        self._stop_event.clear()
        # Tái sử dụng browser từ lần chạy trước (giữ session, tránh logout)
        pw = self._playwright
        br = self._browser
        all_pdf_paths = []
        all_results = []

        try:
            cookie = config['cookie']
            base_dir = config['output_dir']
            master = config['master']
            retail = config['retail']
            auto_print = config['auto_print']
            printer = config['printer']
            test_mode = config['test_mode']
            batch_size = config.get('batch_size', 15)
            pdf_settings = config.get('pdf_settings', 'paper=A4')
            carriers = config['carriers']

            today_str = datetime.now().strftime('%Y-%m-%d')
            out_dir = str(Path(base_dir) / today_str)
            os.makedirs(out_dir, exist_ok=True)

            for carrier, count in carriers:
                if self._stop_event.is_set():
                    self.log_message.emit('warn', 'Đã dừng theo yêu cầu.'); break
                carrier_display = carrier if carrier else 'tất cả'
                count_display = f'{count}' if count > 0 else 'tất cả'
                self.log_message.emit('info', f'📥 Tải PDF [{carrier_display}] ({count_display} đơn)...')

                pdf_paths, pw2, br2 = run_automation(
                    cookie, out_dir, count,
                    lambda m, t='': self.log_message.emit(t, m),
                    lambda s, m: self.state_changed.emit(s, m),
                    self._stop_event,
                    existing_playwright=pw, existing_browser=br,
                    carrier=carrier, test_mode=test_mode)

                if pw2: pw = pw2
                if br2: br = br2
                self._playwright = pw
                self._browser = br

                for p in pdf_paths:
                    if Path(p).exists():
                        self.log_message.emit('info', f'📄 {Path(p).name}')
                        self.result_file.emit(str(p))
                all_pdf_paths.extend(pdf_paths)
                tag = 'ok' if pdf_paths else 'warn'
                self.log_message.emit(tag, f'📥 [{carrier_display}]: đã tải {len(pdf_paths)} file PDF')

                # In shipping label TRƯỚC, rồi mới tính toán (giống Test)
                if auto_print and pdf_paths:
                    self.log_message.emit('info', f'🖨️ [{carrier_display}]: In shipping label...')
                    for p in pdf_paths:
                        name = Path(p).name
                        if not (Path(p).exists() and ('shipping' in name.lower() or 'vận chuyển' in name.lower())):
                            continue
                        try:
                            _print_file(str(p), printer, pdf_settings=pdf_settings, batch_size=batch_size,
                                        log_cb=lambda m, t='': self.log_message.emit(t, m))
                            self.log_message.emit('ok', f'  ✓ Đã gửi in: {Path(p).name}')
                        except Exception as e:
                            self.log_message.emit('err', f'  ✗ Lỗi in {Path(p).name}: {e}')

                carrier_results = []
                if pdf_paths:
                    self.log_message.emit('info', f'📊 Đang tính bill [{carrier_display}]...')
                    carrier_results = run_calculator(
                        pdf_paths, out_dir, master, retail,
                        lambda m, t='': self.log_message.emit(t, m),
                        carrier=carrier)
                    for r in carrier_results:
                        for key, lbl in [('pdf_report', '📄')]:
                            fp = r['files'].get(key)
                            if fp and Path(fp).exists():
                                self.log_message.emit('info', f'{lbl} {Path(fp).name}')
                                self.result_file.emit(f'{lbl} {Path(fp).name}')
                    all_results.extend(carrier_results)

                # In báo cáo sau khi tính toán
                if auto_print and carrier_results:
                    for r in carrier_results:
                        fp = r['files'].get('pdf_report')
                        if fp and Path(fp).exists():
                            try:
                                _print_file(fp, printer, pdf_settings=pdf_settings, batch_size=batch_size,
                                            log_cb=lambda m, t='': self.log_message.emit(t, m))
                                self.log_message.emit('ok', f'  ✓ Đã in báo cáo: {Path(fp).name}')
                            except Exception as e:
                                self.log_message.emit('err', f'  ✗ Lỗi in báo cáo: {e}')

            self.log_message.emit('bold_ok', '🏁 HOÀN THÀNH!')
            self.state_changed.emit('done', f'✅ Hoàn thành lúc {datetime.now().strftime("%H:%M:%S")}')
            self._playwright = pw
            self._browser = br
            self.job_completed.emit({
                'playwright': pw, 'browser': br,
                'pdf_paths': all_pdf_paths, 'results': all_results,
                'output_dir': out_dir,
            })
        except Exception as e:
            self.log_message.emit('err', f'✗ Lỗi: {e}')
            self.state_changed.emit('error', f'✗ {e}')
            self.job_completed.emit({})
        finally:
            self._running = False

    @Slot()
    def stop_job(self): self._stop_event.set()

    @Slot()
    def shutdown(self):
        self._stop_event.set()
        try:
            if self._browser: self._browser.close()
        except: pass
        try:
            if self._playwright: self._playwright.stop()
        except: pass

_warned_sumatra = False
_custom_sumatra_path = ''  # Người dùng có thể chỉ định đường dẫn thủ công trong giao diện


def _wait_print_queue(printer_name, max_jobs=0, timeout=600):
    """Đợi hàng đợi máy in ≤ max_jobs rồi mới gửi job mới (tránh quá tải RAM máy in)."""
    import subprocess as _sp, time as _t
    deadline = _t.time() + timeout
    while _t.time() < deadline:
        result = _sp.run(['powershell', '-NoProfile', '-WindowStyle', 'Hidden', '-Command',
            f"@(Get-PrintJob -PrinterName '{printer_name}' -ErrorAction SilentlyContinue).Count"
        ], capture_output=True, text=True)
        try:
            count = int(result.stdout.strip())
        except ValueError:
            count = 0
        if count <= max_jobs:
            return
        _t.sleep(1)  # kiểm tra mỗi 1 giây


def _ensure_sumatra_settings():
    """Tạo/ghi đè file cài đặt SumatraPDF để tắt Welcome page."""
    try:
        settings_dir = Path(os.environ.get('APPDATA', '')) / 'SumatraPDF'
        settings_file = settings_dir / 'SumatraPDF-settings.txt'
        settings_dir.mkdir(parents=True, exist_ok=True)
        # Đọc nội dung cũ (nếu có), chỉ ghi đè dòng ShowStartPage
        content = ''
        if settings_file.exists():
            content = settings_file.read_text()
        # Luôn đảm bảo ShowStartPage = false
        if 'ShowStartPage' in content:
            content = _re.sub(r'^ShowStartPage\s*=.*$', 'ShowStartPage = false', content, flags=_re.MULTILINE)
        else:
            content = content.rstrip('\n') + '\nShowStartPage = false\n'
        if 'RememberOpenedFiles' not in content:
            content = content.rstrip('\n') + '\nRememberOpenedFiles = false\n'
        settings_file.write_text(content)
    except Exception:
        pass  # Không ảnh hưởng đến luồng in chính

def _print_file(file_path, printer_name, pdf_settings='paper=A4', log_cb=None, batch_size=15):
    import subprocess, os as _os
    fp = str(file_path)
    if fp.lower().endswith('.pdf'):
        sumatra_exe = None
        # Ưu tiên: 0) Đường dẫn tùy chỉnh từ giao diện  1) Cạnh exe  2) %LOCALAPPDATA%  3) Program Files
        sumatra_paths = []
        if _custom_sumatra_path and Path(_custom_sumatra_path).exists():
            sumatra_paths.append(_custom_sumatra_path)
        sumatra_paths += [
            _os.path.join(_os.path.dirname(sys.executable), 'SumatraPDF.exe'),
            _os.path.join(_os.environ.get('LOCALAPPDATA', ''), 'SumatraPDF', 'SumatraPDF.exe'),
            r'C:\Program Files\SumatraPDF\SumatraPDF.exe',
            r'C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe',
        ]
        for sp in sumatra_paths:
            if Path(sp).exists(): sumatra_exe = sp; break
        print_path = fp
        temp_merged = None
        # Chỉ merge 2-up file shipping label, không merge file báo cáo
        fname_lower = Path(fp).name.lower()
        if 'shipping' in fname_lower or 'vận chuyển' in fname_lower:
            try:
                temp_merged = _merge_pdf_2up(fp)
                if temp_merged: print_path = temp_merged
            except: pass

        if sumatra_exe:
            _ensure_sumatra_settings()
            # Chia batch nếu file > batch_size tờ (tránh máy in hết RAM)
            try:
                from pypdf import PdfReader as _PdfReader, PdfWriter as _PdfWriter
                reader = _PdfReader(print_path)
                total_pages = len(reader.pages)
            except Exception:
                reader = None
                total_pages = 0

            if reader and total_pages > batch_size:
                total_batches = (total_pages + batch_size - 1) // batch_size
                if log_cb: log_cb(f'  📦 Chia {total_pages} tờ → {total_batches} batch ({batch_size} tờ/batch)', 'info')
                batch_num = 0
                for start in range(0, total_pages, batch_size):
                    batch_num += 1
                    end = min(start + batch_size, total_pages)
                    if log_cb: log_cb(f'  🖨️ Batch {batch_num}/{total_batches} (tờ {start+1}-{end})...', 'info')
                    batch_writer = _PdfWriter()
                    for i in range(start, end):
                        batch_writer.add_page(reader.pages[i])
                    batch_path = print_path + f'.batch{batch_num}.pdf'
                    with open(batch_path, 'wb') as bf:
                        batch_writer.write(bf)

                    # Đợi queue trống rồi gửi batch tiếp (in liền mạch, không ngắt quãng)
                    _wait_print_queue(printer_name, max_jobs=0)

                    cmd = [sumatra_exe, '-print-to', printer_name, '-exit-when-done', batch_path]
                    if pdf_settings:
                        cmd += ['-print-settings', pdf_settings]
                    result = subprocess.run(cmd, check=False, timeout=600)
                    if result.returncode != 0:
                        raise RuntimeError(f'SumatraPDF batch {batch_num} exit code: {result.returncode}')

                    # Dọn file batch tạm
                    try: _os.remove(batch_path)
                    except: pass

                # Đợi batch cuối in xong
                _wait_print_queue(printer_name)
            else:
                _wait_print_queue(printer_name)
                cmd = [sumatra_exe, '-print-to', printer_name, '-exit-when-done', print_path]
                if pdf_settings: cmd += ['-print-settings', pdf_settings]
                result = subprocess.run(cmd, check=False, timeout=600)
                if result.returncode != 0:
                    raise RuntimeError(f'SumatraPDF exit code: {result.returncode}')
        else:
            global _warned_sumatra
            if not _warned_sumatra:
                _warned_sumatra = True
                # Log rõ ràng để user biết cần cài SumatraPDF
                print(f'[TTS_Bill] ⚠ Không tìm thấy SumatraPDF.exe. '
                      f'Đặt file vào thư mục chứa TTS_Bill.exe hoặc cài tại C:\\Program Files\\SumatraPDF.')
            raise RuntimeError(
                'Không tìm thấy SumatraPDF.exe. '
                'Đặt SumatraPDF.exe vào thư mục portable hoặc cài đặt SumatraPDF.'
            )
        if temp_merged:
            def _cleanup(p=temp_merged):
                import time; time.sleep(5)
                try: _os.remove(p)
                except: pass
            threading.Timer(5, _cleanup).start()
        return

    if fp.lower().endswith('.xlsx') or fp.lower().endswith('.xls'):
        try:
            import pythoncom, win32com.client
            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            workbook = excel.Workbooks.Open(_os.path.abspath(fp))
            workbook.PrintOut(ActivePrinter=printer_name)
            workbook.Close(False); excel.Quit(); pythoncom.CoUninitialize()
            return
        except Exception:
            try: pythoncom.CoUninitialize()
            except: pass
        # Nếu COM thất bại (không có Excel) → báo lỗi rõ ràng
        raise RuntimeError(
            'Không thể in file Excel. Máy cần cài Microsoft Excel.'
        )

def _merge_pdf_2up(pdf_path):
    import os as _os
    try: from pypdf import PdfReader, PdfWriter, PageObject, Transformation
    except ImportError: return None
    try:
        reader = PdfReader(pdf_path)
        if len(reader.pages) < 2: return None
        canvas_w, canvas_h = 842, 595
        margin_left, margin_top, gap, scale_factor = 14, 28, -14, 0.70
        writer = PdfWriter()
        avail_w = canvas_w - 2*margin_left - gap
        half_w = avail_w / 2
        for pair_start in range(0, len(reader.pages), 2):
            pair = reader.pages[pair_start:pair_start+2]
            canvas = PageObject.create_blank_page(width=canvas_w, height=canvas_h)
            for i, page in enumerate(pair):
                pw = float(page.mediabox.width); ph = float(page.mediabox.height)
                scale = min(half_w / pw, (canvas_h - 2*margin_top) / ph) * scale_factor
                sw, sh = pw * scale, ph * scale
                tx = margin_left + i * (half_w + gap); ty = canvas_h - margin_top - sh
                canvas.merge_transformed_page(page, Transformation().scale(scale).translate(tx / scale, ty / scale))
            writer.add_page(canvas)
        # Nén để giảm dung lượng (tránh file 2up nặng hơn file gốc 9 lần)
        for page in writer.pages:
            page.compress_content_streams()
        temp_path = pdf_path + '.2up.pdf'
        with open(temp_path, 'wb') as f: writer.write(f)
        return temp_path
    except: return None

def _get_printers() -> list:
    printers = []
    try:
        import win32print
        for info in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS, None, 1):
            name = info[2].strip() if info[2] else ''
            if name: printers.append(name)
    except: pass
    if not printers:
        try:
            import subprocess
            result = subprocess.run(['powershell', '-Command', "Get-Printer | Select-Object -ExpandProperty Name | Where-Object { $_ -notlike '*Microsoft*' -and $_ -notlike '*Fax*' -and $_ -notlike '*OneNote*' -and $_ -notlike '*XPS*' }"], capture_output=True, text=True, timeout=10)
            for line in result.stdout.strip().split('\n'):
                name = line.strip()
                if name and name not in printers: printers.append(name)
        except: pass
    return printers

# ═══════════════════════════════════════════════════════════════
# WIDGETS
# ═══════════════════════════════════════════════════════════════
class FileRowWidget(QWidget):
    path_changed = Signal(str)

    def __init__(self, label_text: str, file_filter: str = '', is_dir: bool = False, parent=None):
        super().__init__(parent)
        self._real_path = ''
        self._filter = file_filter
        self._is_dir = is_dir

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)

        row = QHBoxLayout()
        row.setSpacing(6)

        lbl = QLabel(label_text)
        lbl.setFixedWidth(130)
        lbl.setStyleSheet('font-weight: 600; color: #1E293B; font-size: 10pt;')
        row.addWidget(lbl)

        self.path_edit = QLineEdit()
        self.path_edit.setReadOnly(True)
        self.path_edit.setObjectName('pathEdit')
        self.path_edit.setPlaceholderText('Chưa chọn...')
        row.addWidget(self.path_edit, 1)

        btn = QPushButton('Chọn')
        btn.setObjectName('browseBtn')
        btn.setCursor(Qt.PointingHandCursor)
        btn.clicked.connect(self._browse)
        row.addWidget(btn)

        layout.addLayout(row)

        self.status_label = QLabel('❌ Chưa chọn')
        self.status_label.setObjectName('statusLabel')
        self.status_label.setStyleSheet('font-size: 8.5pt; color: #DC2626; padding-left: 136px;')
        layout.addWidget(self.status_label)

    def _browse(self):
        if self._is_dir:
            p = QFileDialog.getExistingDirectory(self, 'Chọn thư mục')
            if p:
                self._real_path = p
                self.path_edit.setText(p)
                self._update_status_dir()
                self.path_changed.emit(p)
        else:
            p, _ = QFileDialog.getOpenFileName(self, 'Chọn file', '', self._filter)
            if p:
                self._real_path = p
                self.path_edit.setText(p)
                self.path_changed.emit(p)

    def set_path(self, path: str):
        self._real_path = path
        self.path_edit.setText(path)

    def get_real_path(self) -> str:
        return self._real_path

    def _update_status_dir(self):
        if self._real_path and Path(self._real_path).exists():
            self.status_label.setText('✅ Đã cập nhật thư mục lưu')
            self.status_label.setStyleSheet('font-size: 8.5pt; color: #059669; padding-left: 136px;')
        else:
            self.status_label.setText('❌ Chưa chọn')
            self.status_label.setStyleSheet('font-size: 8.5pt; color: #DC2626; padding-left: 136px;')

    def update_cookie_status(self):
        p = Path(self._real_path) if self._real_path else Path('')
        if self._real_path and p.exists():
            try:
                d = json.loads(p.read_text(encoding='utf-8'))
                cs = d.get('cookies', d)
                n = len(cs) if isinstance(cs, list) else 0
                self.status_label.setText(f'✅ Đã nạp {n} cookies')
                self.status_label.setStyleSheet('font-size: 8.5pt; color: #059669; padding-left: 136px;')
            except Exception:
                self.status_label.setText('⚠ File bị lỗi hoặc sai định dạng')
                self.status_label.setStyleSheet('font-size: 8.5pt; color: #D97706; padding-left: 136px;')
        else:
            self.status_label.setText('❌ Chưa chọn file hợp lệ')
            self.status_label.setStyleSheet('font-size: 8.5pt; color: #DC2626; padding-left: 136px;')

    def update_excel_status(self):
        p = Path(self._real_path) if self._real_path else Path('')
        if self._real_path and p.exists() and p.suffix.lower() == '.xlsx':
            try:
                from openpyxl import load_workbook
                wb = load_workbook(str(p), data_only=True, read_only=True)
                sheet = wb.active
                n = sheet.max_row - 1 if sheet.max_row else 0
                wb.close()
                self.status_label.setText(f'✅ Đã nạp {n} SKU')
                self.status_label.setStyleSheet('font-size: 8.5pt; color: #059669; padding-left: 136px;')
            except Exception:
                self.status_label.setText('⚠ File bị lỗi')
                self.status_label.setStyleSheet('font-size: 8.5pt; color: #D97706; padding-left: 136px;')
        else:
            self.status_label.setText('❌ Chưa chọn file hợp lệ')
            self.status_label.setStyleSheet('font-size: 8.5pt; color: #DC2626; padding-left: 136px;')


# ============================================================
# MAIN APP — QMainWindow with PySide6 GUI (SaaS Light Theme)
# ============================================================
class App(QMainWindow):
    trigger_job = Signal(dict)
    _test_log = Signal(str, str)  # signal cho test tab (thread-safe)

    TAG_COLORS = {
        "ts": "#64748B", "ok": "#10B981", "err": "#F87171",
        "warn": "#FBBF24", "info": "#60A5FA", "batch": "#C084FC",
        "result": "#F472B6", "dim": "#64748B", "bold_ok": "#10B981",
    }

    def __init__(self):
        super().__init__()
        self.setWindowTitle("TikTokPrint")
        self.resize(850, 850)
        self.setMinimumSize(600, 700)

        self._cookie_real = str(DEFAULT_COOKIE) if DEFAULT_COOKIE.exists() else ""
        local_master = BASE_DIR / "mã combo.xlsx"
        if local_master.exists(): self._master_real = str(local_master)
        elif MASTER_DEFAULT.exists(): self._master_real = str(MASTER_DEFAULT)
        else: self._master_real = ""
        local_retail = BASE_DIR / "sp bán lẻ.xlsx"
        if local_retail.exists(): self._retail_real = str(local_retail)
        elif RETAIL_DEFAULT.exists(): self._retail_real = str(RETAIL_DEFAULT)
        else: self._retail_real = ""

        self.running = False
        self.scheduler_active = False
        self.result_files = []
        self._stop_event = threading.Event()

        self._sched_mode = "once"
        self._sched_interval_hours = 1
        self._sched_daily_times = []
        self._sched_next_run = None
        self._sched_last_run = None

        self._worker_thread = QThread()
        self._worker = AutomationWorker()
        self._worker.moveToThread(self._worker_thread)
        self._worker.log_message.connect(self._on_log_message)
        self._worker.state_changed.connect(self._on_state_changed)
        self._worker.job_completed.connect(self._on_job_completed)
        self._worker.result_file.connect(self._add_result)
        self.trigger_job.connect(self._worker.start_job)
        self._test_log.connect(self._log_html)  # test tab log (thread-safe)
        self._worker_thread.start()

        self._sched_timer = QTimer(self)
        self._sched_timer.setInterval(1000)
        self._sched_timer.timeout.connect(self._check_schedule)

        self._apply_stylesheet()
        self._build_ui()

        # Set global SumatraPDF path nếu đã detect được
        if self.sumatra_path_edit.text():
            self._on_sumatra_path_changed(self.sumatra_path_edit.text())

        self._update_cookie_status()
        self._update_master_status()
        self._update_retail_status()
        out_dir = self.output_row.get_real_path()
        if out_dir:
            os.makedirs(out_dir, exist_ok=True)
        self._load_preview_data()

    # ═══════════════════════════════════════════════════════
    # QSS — FIXED BLACK BACKGROUND BUG
    # ═══════════════════════════════════════════════════════
    def _apply_stylesheet(self):
        qss = """
            QMainWindow { background: #F8FAFC; }
            
            /* Fix Black Background Bug on ScrollArea and StackedWidget */
            QWidget#scrollContent { background: #F8FAFC; }
            QScrollArea, QStackedWidget { background: #F8FAFC; border: none; }
            
            /* Header */
            QWidget#headerBar { background: #065F46; }
            QWidget#headerBar QLabel#headerTitle { color: #FFFFFF; font-size: 20px; font-weight: bold; }
            QWidget#headerBar QLabel#headerSubtitle { color: #A7F3D0; font-size: 12px; margin-top: 2px;}

            /* Tabs Navigation */
            QWidget#tabBar { background: #FFFFFF; border-bottom: 1px solid #E2E8F0; }
            QPushButton#tabBtn { background: transparent; color: #64748B; font-weight: bold; font-size: 14px; padding: 14px 24px; border: none; border-bottom: 3px solid transparent; }
            QPushButton#tabBtn:hover { color: #1E293B; border-bottom: 3px solid #CBD5E1; }
            QPushButton#tabBtn[active="true"] { color: #059669; border-bottom: 3px solid #059669; }

            /* GroupBox */
            QGroupBox { 
                background: #FFFFFF; 
                border: 1px solid #E2E8F0; 
                border-radius: 12px; 
                margin-top: 24px; 
                padding-top: 36px; 
                padding-left: 20px; 
                padding-right: 20px; 
                padding-bottom: 20px; 
                font-weight: bold; 
                font-size: 15px; 
                color: #1E293B; 
            }
            QGroupBox::title { subcontrol-origin: margin; subcontrol-position: top left; left: 20px; top: 0px; color: #1E293B; font-size: 15px; font-weight: bold; }

            /* Inputs */
            QLineEdit, QSpinBox, QComboBox { background: #F1F5F9; border: 1px solid #E2E8F0; border-radius: 6px; padding: 8px 12px; color: #1E293B; font-size: 14px; }
            QLineEdit:focus, QSpinBox:focus, QComboBox:focus { border-color: #059669; background: #FFFFFF; }
            
            /* Search Box specifically */
            QLineEdit#searchBox { border-radius: 16px; padding: 6px 14px; font-size: 13px; }

            QComboBox::drop-down { border: none; width: 24px; }
            QComboBox QAbstractItemView { background: #FFFFFF; border: 1px solid #E2E8F0; selection-background-color: #D1FAE5; selection-color: #1E293B; }

            /* Regular Buttons */
            QPushButton#browseBtn, QPushButton#refreshBtn { background: #FFFFFF; border: 1px solid #CBD5E1; border-radius: 6px; padding: 8px 16px; color: #1E293B; font-weight: bold; font-size: 14px; }
            QPushButton#browseBtn:hover, QPushButton#refreshBtn:hover { background: #F8FAFC; border-color: #94A3B8; }

            /* Action Buttons */
            QPushButton#schedBtn { background: #059669; color: #FFFFFF; border: none; border-radius: 8px; padding: 12px; font-weight: bold; font-size: 16px; }
            QPushButton#schedBtn:hover { background: #047857; }
            QPushButton#schedBtn:disabled { background: #A7F3D0; }

            QPushButton#stopBtn { background: #FEF2F2; color: #DC2626; border: 1px solid #FECACA; border-radius: 8px; padding: 12px; font-weight: bold; font-size: 14px; }
            QPushButton#stopBtn:hover { background: #FEE2E2; }
            QPushButton#stopBtn:disabled { background: #F8FAFC; color: #9CA3AF; border-color: #E2E8F0; }

            /* Small Icon Buttons */
            QPushButton#smallBtn { background: transparent; border: none; padding: 6px; color: #64748B; font-size: 12px; font-weight: bold; }
            QPushButton#smallBtn:hover { background: #F1F5F9; color: #1E293B; border-radius: 4px; }

            /* Preview Tabs */
            QPushButton#previewTabBtn { background: transparent; border: 1px solid transparent; color: #64748B; padding: 6px 12px; border-radius: 4px; font-size: 13px; font-weight: bold; }
            QPushButton#previewTabBtn[active="true"] { background: #F1F5F9; border: 1px solid #E2E8F0; color: #1E293B; }

            /* Tables */
            QTableWidget { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; gridline-color: #E2E8F0; font-size: 13px; color: #1E293B; }
            QTableWidget::item { padding: 4px; }
            QTableWidget::item:selected { background: #F8FAFC; color: #1E293B; }
            QTableWidget QLineEdit { padding: 1px 2px; background: #FFFFFF; border: 2px solid #059669; border-radius: 2px; color: #1E293B; font-size: 13px; }
            QHeaderView::section { background: #F1F5F9; color: #1E293B; font-weight: bold; padding: 10px 12px; border: none; border-right: 1px solid #E2E8F0; border-bottom: 1px solid #E2E8F0; font-size: 13px; }

            /* Checkboxes & Radio */
            QCheckBox, QRadioButton { color: #1E293B; font-size: 10pt; spacing: 8px; }
            QCheckBox::indicator, QRadioButton::indicator { width: 16px; height: 16px; }

            /* Log */
            QTextEdit#logView { background: #0F172A; color: #F8FAFC; border: 1px solid #E2E8F0; border-radius: 8px; font-family: "Consolas", monospace; font-size: 13px; padding: 12px; selection-background-color: #1E3A5F; selection-color: #FFFFFF; }

            /* Result List */
            QListWidget#resultList { background: #FFFFFF; border: 1px solid #E2E8F0; border-radius: 8px; padding: 4px; color: #1E293B; font-size: 13px; }
            QListWidget#resultList::item { padding: 8px; border-bottom: 1px solid #F1F5F9; }
            QListWidget#resultList::item:selected { background: #D1FAE5; color: #047857; font-weight: bold; border-radius: 4px; }

            /* Scrollbars */
            QScrollBar:vertical { background: transparent; width: 8px; }
            QScrollBar::handle:vertical { background: #CBD5E1; border-radius: 4px; min-height: 30px; }
            QScrollBar::handle:vertical:hover { background: #94A3B8; }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }

            /* Bottom Bar */
            QWidget#bottomBar { background: #FFFFFF; border-top: 1px solid #E2E8F0; }
            QLabel#statusLabel { font-weight: bold; font-size: 14px; padding: 4px 0; }
            QFrame#divider { background: #E2E8F0; max-height: 1px; margin: 12px 0; }
        """
        self.setStyleSheet(qss)
        font = QFont("Segoe UI", 10)
        self.setFont(font)

    # ═══════════════════════════════════════════════════════
    # BUILD UI
    # ═══════════════════════════════════════════════════════
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        main_layout.addWidget(self._build_header())
        main_layout.addWidget(self._build_tab_bar())

        self.content_stack = QStackedWidget()
        self.content_stack.addWidget(self._build_files_tab())
        self.content_stack.addWidget(self._build_print_tab())
        self.content_stack.addWidget(self._build_schedule_tab())
        self.content_stack.addWidget(self._build_log_tab())
        self.content_stack.addWidget(self._build_test_tab())
        main_layout.addWidget(self.content_stack, 1)

        main_layout.addWidget(self._build_bottom_bar())

    def _build_header(self):
        h = QWidget()
        h.setObjectName("headerBar")
        h.setFixedHeight(70)
        layout = QVBoxLayout(h)
        layout.setContentsMargins(24, 12, 24, 12)
        layout.setSpacing(2)

        title = QLabel("TikTokPrint")
        title.setObjectName("headerTitle")
        layout.addWidget(title)

        subtitle = QLabel("Automation & Bill Calculate")
        subtitle.setObjectName("headerSubtitle")
        layout.addWidget(subtitle)

        return h

    def _build_tab_bar(self):
        bar = QWidget()
        bar.setObjectName("tabBar")
        layout = QHBoxLayout(bar)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.tab_buttons = []
        self.tab_button_group = QButtonGroup(self)
        self.tab_button_group.setExclusive(True)

        tab_labels = [
            "📦 Tệp dữ liệu",
            "🖨️ Cấu hình In",
            "⏰ Lịch trình",
            "📋 Nhật ký & Kết quả",
            "🧪 Test",
        ]

        for i, label in enumerate(tab_labels):
            btn = QPushButton(label)
            btn.setObjectName("tabBtn")
            btn.setCheckable(True)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setProperty("tabIndex", i)
            if i == 0:
                btn.setChecked(True)
                btn.setProperty("active", True)
                btn.style().unpolish(btn)
                btn.style().polish(btn)
            self.tab_button_group.addButton(btn, i)
            self.tab_buttons.append(btn)
            layout.addWidget(btn)

        layout.addStretch()
        self.tab_button_group.idClicked.connect(self._on_tab_changed)
        return bar

    def _on_tab_changed(self, idx: int):
        self.content_stack.setCurrentIndex(idx)
        for i, btn in enumerate(self.tab_buttons):
            is_active = (i == idx)
            btn.setProperty("active", is_active)
            btn.style().unpolish(btn)
            btn.style().polish(btn)

    # ═══════════════════════════════════════════════════════
    # TAB 0: TỆP DỮ LIỆU & BẢNG PREVIEW
    # ═══════════════════════════════════════════════════════
    def _build_files_tab(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        
        w = QWidget()
        w.setObjectName("scrollContent")
        layout = QVBoxLayout(w)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        # ── Group 1: Khai báo đường dẫn ──
        gb1 = QGroupBox("Khai báo đường dẫn")
        gb1_layout = QVBoxLayout(gb1)
        gb1_layout.setSpacing(12)
        gb1_layout.setContentsMargins(0, 0, 0, 0) # Use QSS paddings

        self.cookie_row = FileRowWidget("🍪 Cookie (JSON)", "JSON Files (*.json)")
        self.cookie_row.set_path(self._cookie_real)
        self.cookie_row.path_changed.connect(self._on_cookie_changed)
        gb1_layout.addWidget(self.cookie_row)

        self.master_row = FileRowWidget("📦 Master (Combo)", "Excel Files (*.xlsx)")
        self.master_row.set_path(self._master_real)
        self.master_row.path_changed.connect(self._on_master_changed)
        gb1_layout.addWidget(self.master_row)

        self.retail_row = FileRowWidget("🛍 Bán lẻ", "Excel Files (*.xlsx)")
        self.retail_row.set_path(self._retail_real)
        self.retail_row.path_changed.connect(self._on_retail_changed)
        gb1_layout.addWidget(self.retail_row)

        self.output_row = FileRowWidget("📂 Thư mục lưu", is_dir=True)
        self.output_row.set_path(str(BASE_DIR / "outputs"))
        self.output_row._update_status_dir()
        self.output_row.path_changed.connect(self._on_output_dir_changed)
        gb1_layout.addWidget(self.output_row)

        layout.addWidget(gb1)

        # ── Group 2: Xem trước dữ liệu Excel & Ô Tìm Kiếm ──
        gb2 = QGroupBox("Xem trước dữ liệu Excel (Hiển thị tối đa 5000 dòng)")
        gb2_layout = QVBoxLayout(gb2)
        gb2_layout.setSpacing(12)
        gb2_layout.setContentsMargins(0, 0, 0, 0)

        toolbar = QHBoxLayout()
        toolbar.setSpacing(8)

        self.preview_btn_group = QButtonGroup(self)
        self.preview_btn_group.setExclusive(True)

        self.preview_master_btn = QPushButton("Bảng Master (Combo)")
        self.preview_master_btn.setObjectName("previewTabBtn")
        self.preview_master_btn.setCheckable(True)
        self.preview_master_btn.setChecked(True)
        self.preview_master_btn.setProperty("active", True)
        self.preview_master_btn.setCursor(Qt.PointingHandCursor)
        self.preview_btn_group.addButton(self.preview_master_btn, 0)
        toolbar.addWidget(self.preview_master_btn)

        self.preview_retail_btn = QPushButton("Bảng Bán lẻ")
        self.preview_retail_btn.setObjectName("previewTabBtn")
        self.preview_retail_btn.setCheckable(True)
        self.preview_retail_btn.setCursor(Qt.PointingHandCursor)
        self.preview_btn_group.addButton(self.preview_retail_btn, 1)
        toolbar.addWidget(self.preview_retail_btn)
        self.preview_btn_group.idClicked.connect(self._on_preview_tab_changed)

        toolbar.addStretch()

        self.search_input = QLineEdit()
        self.search_input.setObjectName("searchBox")
        self.search_input.setPlaceholderText("🔍 Tìm kiếm SKU, tên sản phẩm...")
        self.search_input.setFixedWidth(250)
        self.search_input.textChanged.connect(self._filter_preview_table)
        toolbar.addWidget(self.search_input)

        gb2_layout.addLayout(toolbar)

        self.master_table = QTableWidget()
        self.master_table.setEditTriggers(QTableWidget.DoubleClicked)
        self.master_table.setMinimumHeight(450) # Đảm bảo bảng luôn mở rộng thoải mái để xem
        self.master_table.cellChanged.connect(self._on_table_cell_changed)
        gb2_layout.addWidget(self.master_table, 1)

        self.retail_table = QTableWidget()
        self.retail_table.setEditTriggers(QTableWidget.DoubleClicked)
        self.retail_table.setMinimumHeight(450) # Đảm bảo bảng luôn mở rộng thoải mái để xem
        self.retail_table.cellChanged.connect(self._on_table_cell_changed)
        self.retail_table.hide()
        gb2_layout.addWidget(self.retail_table, 1)

        layout.addWidget(gb2, 1)
        scroll.setWidget(w)
        return scroll

    # ═══════════════════════════════════════════════════════
    # TAB 1: CẤU HÌNH IN 
    # ═══════════════════════════════════════════════════════
    def _build_print_tab(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        w = QWidget()
        w.setObjectName("scrollContent")
        layout = QVBoxLayout(w)
        layout.setContentsMargins(24, 24, 24, 24)

        gb = QGroupBox("Thiết lập tải và in tự động")
        gb_layout = QVBoxLayout(gb)
        gb_layout.setSpacing(12)
        gb_layout.setContentsMargins(0, 0, 0, 0)

        # ── Carrier grid: 2 cột ──
        grid = QGridLayout()
        grid.setHorizontalSpacing(32)
        grid.setVerticalSpacing(16)

        self.carrier_spinboxes = {}
        self.carrier_checkboxes = {}
        carriers = [
            ("J&T Express:", "jt"), ("GHN:", "ghn"),
            ("VietNam Post:", "vnp"), ("Best Express:", "best"),
            ("Viettel Post:", "viettel"), ("J&T Cargo VN:", "jtc"),
        ]
        for idx, (label, key) in enumerate(carriers):
            row, col = idx % 3, (idx // 3)

            box = QWidget()
            box_ly = QHBoxLayout(box)
            box_ly.setContentsMargins(0,0,0,0)

            cb = QCheckBox(label.replace(":", ""))
            cb.setChecked(True)  # mặc định bật tất cả
            cb.setStyleSheet("font-weight: 500;")
            cb.setFixedWidth(120)
            box_ly.addWidget(cb)
            self.carrier_checkboxes[key] = cb

            sb = QSpinBox()
            sb.setRange(0, 9999)
            sb.setValue(0)
            sb.setFixedWidth(65)
            sb.setAlignment(Qt.AlignRight)
            sb.setToolTip("0 = tất cả đơn của hãng này")
            box_ly.addWidget(sb)

            hint = QLabel("đơn  (0 = tất cả)")
            hint.setStyleSheet("color: #64748B; font-size: 11px;")
            box_ly.addWidget(hint)

            # Checkbox OFF → disable spinbox (bỏ qua hãng này)
            cb.toggled.connect(sb.setEnabled)

            box_ly.addStretch()

            grid.addWidget(box, row, col)
            self.carrier_spinboxes[key] = sb

        gb_layout.addLayout(grid)

        # Divider
        div = QFrame()
        div.setObjectName("divider")
        div.setFrameShape(QFrame.HLine)
        gb_layout.addWidget(div)

        # ── In tự động row ──
        print_row = QHBoxLayout()
        print_row.setSpacing(12)

        self.auto_print_cb = QCheckBox("🖨️ In tự động ra máy in")
        self.auto_print_cb.setStyleSheet("font-weight: bold; color: #059669;")
        print_row.addWidget(self.auto_print_cb)

        printers_list = _get_printers()
        self.printer_combo = QComboBox()
        self.printer_combo.addItems(printers_list)
        self.printer_combo.setMinimumWidth(180)
        if printers_list:
            self.printer_combo.setCurrentIndex(0)
        print_row.addWidget(self.printer_combo, 1)

        refresh_btn = QPushButton("↻ Cập nhật")
        refresh_btn.setObjectName("refreshBtn")
        refresh_btn.setCursor(Qt.PointingHandCursor)
        refresh_btn.clicked.connect(self._refresh_printers)
        print_row.addWidget(refresh_btn)

        gb_layout.addLayout(print_row)

        # ── Paper settings row ──
        paper_row = QHBoxLayout()
        paper_row.setSpacing(12)
        paper_row.setContentsMargins(0, 8, 0, 0)

        paper_row.addWidget(QLabel("In mặt:"))
        self.duplex_combo = QComboBox()
        self.duplex_combo.addItems(["simplex", "longedge", "shortedge"])
        self.duplex_combo.setFixedWidth(100)
        paper_row.addWidget(self.duplex_combo)

        paper_row.addWidget(QLabel("  Batch in (tờ/lần):"))
        self.batch_size_spin = QSpinBox()
        self.batch_size_spin.setRange(1, 100)
        self.batch_size_spin.setValue(15)
        self.batch_size_spin.setFixedWidth(60)
        self.batch_size_spin.setToolTip("Số tờ in mỗi lần, tránh máy in quá tải")
        paper_row.addWidget(self.batch_size_spin)

        paper_row.addStretch()
        gb_layout.addLayout(paper_row)

        # ── SumatraPDF custom path row ──
        sumatra_row = QHBoxLayout()
        sumatra_row.setSpacing(8)
        sumatra_row.setContentsMargins(0, 4, 0, 0)
        sumatra_row.addWidget(QLabel("SumatraPDF:"))

        # Tự động dò tìm SumatraPDF
        import os as _os_detect
        _detected = ''
        for _sp in [
            _os_detect.path.join(_os_detect.path.dirname(sys.executable), 'SumatraPDF.exe'),
            _os_detect.path.join(_os_detect.environ.get('LOCALAPPDATA', ''), 'SumatraPDF', 'SumatraPDF.exe'),
            r'C:\Program Files\SumatraPDF\SumatraPDF.exe',
            r'C:\Program Files (x86)\SumatraPDF\SumatraPDF.exe',
        ]:
            if Path(_sp).exists():
                _detected = _sp; break

        self.sumatra_path_edit = QLineEdit()
        self.sumatra_path_edit.setPlaceholderText("Không tìm thấy — chọn thủ công...")
        self.sumatra_path_edit.setMinimumWidth(280)
        if _detected:
            self.sumatra_path_edit.setText(_detected)
            self.sumatra_path_edit.setStyleSheet("color: #059669;")  # xanh lá = đã tìm thấy
        self.sumatra_path_edit.textChanged.connect(self._on_sumatra_path_changed)
        sumatra_row.addWidget(self.sumatra_path_edit, 1)
        browse_sumatra_btn = QPushButton("Duyệt...")
        browse_sumatra_btn.setFixedWidth(70)
        browse_sumatra_btn.setCursor(Qt.PointingHandCursor)
        browse_sumatra_btn.clicked.connect(self._browse_sumatra_path)
        sumatra_row.addWidget(browse_sumatra_btn)
        gb_layout.addLayout(sumatra_row)

        # ── Test mode ──
        self.test_mode_cb = QCheckBox("🧪 Bật chế độ Test Mode (Chỉ tải danh sách đơn, KHÔNG thao tác in)")
        self.test_mode_cb.setChecked(True)
        self.test_mode_cb.setStyleSheet("color: #D97706; font-weight: 600; font-size: 14px; margin-top: 12px;")
        gb_layout.addWidget(self.test_mode_cb)

        layout.addWidget(gb)
        layout.addStretch()
        scroll.setWidget(w)
        return scroll

    # ═══════════════════════════════════════════════════════
    # TAB 2: LỊCH TRÌNH
    # ═══════════════════════════════════════════════════════
    def _build_schedule_tab(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        w = QWidget()
        w.setObjectName("scrollContent")
        layout = QVBoxLayout(w)
        layout.setContentsMargins(24, 24, 24, 24)

        gb = QGroupBox("Thiết lập bộ đếm thời gian")
        gb_layout = QVBoxLayout(gb)
        gb_layout.setSpacing(16)
        gb_layout.setContentsMargins(0, 0, 0, 0)

        self.sched_button_group = QButtonGroup(self)

        self.once_rb = QRadioButton("▶ Chạy ngay lập tức 1 lần duy nhất")
        self.once_rb.setChecked(True)
        self.once_rb.setProperty("mode", "once")
        self.sched_button_group.addButton(self.once_rb)
        gb_layout.addWidget(self.once_rb)

        self.interval_rb = QRadioButton("🔄 Lặp lại tự động mỗi N giờ")
        self.interval_rb.setProperty("mode", "interval")
        self.sched_button_group.addButton(self.interval_rb)
        gb_layout.addWidget(self.interval_rb)

        # Interval sub-panel
        self.interval_panel = QWidget()
        ip_layout = QHBoxLayout(self.interval_panel)
        ip_layout.setContentsMargins(32, 0, 0, 0)
        ip_layout.setSpacing(8)
        ip_layout.addWidget(QLabel("Khởi chạy lại mỗi:"))
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 24)
        self.interval_spin.setValue(1)
        self.interval_spin.setFixedWidth(80)
        ip_layout.addWidget(self.interval_spin)
        lbl_h = QLabel("giờ")
        lbl_h.setStyleSheet("color: #64748B;")
        ip_layout.addWidget(lbl_h)
        ip_layout.addStretch()
        self.interval_panel.hide()
        gb_layout.addWidget(self.interval_panel)

        self.daily_rb = QRadioButton("🕒 Chạy vào các khung giờ cố định trong ngày")
        self.daily_rb.setProperty("mode", "daily")
        self.sched_button_group.addButton(self.daily_rb)
        gb_layout.addWidget(self.daily_rb)

        # Daily sub-panel
        self.daily_panel = QWidget()
        dp_layout = QHBoxLayout(self.daily_panel)
        dp_layout.setContentsMargins(32, 0, 0, 0)
        dp_layout.setSpacing(8)
        dp_layout.addWidget(QLabel("Khung giờ:"))
        self.daily_times_edit = QLineEdit("08:00, 14:00, 20:00")
        self.daily_times_edit.setFixedWidth(200)
        dp_layout.addWidget(self.daily_times_edit)
        hint = QLabel("(Ngăn cách bằng dấu phẩy)")
        hint.setStyleSheet("color: #64748B; font-size: 13px;")
        dp_layout.addWidget(hint)
        dp_layout.addStretch()
        self.daily_panel.hide()
        gb_layout.addWidget(self.daily_panel)

        self.sched_button_group.buttonClicked.connect(self._on_schedule_mode_changed)

        layout.addWidget(gb)
        layout.addStretch()
        scroll.setWidget(w)
        return scroll

    # ═══════════════════════════════════════════════════════
    # TAB 3: NHẬT KÝ & KẾT QUẢ
    # ═══════════════════════════════════════════════════════
    def _build_log_tab(self):
        w = QWidget()
        w.setObjectName("scrollContent")
        layout = QVBoxLayout(w)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(24)

        # ── Log group (flex grow) ──
        log_gb = QGroupBox("Nhật ký hệ thống (Log)")
        log_layout = QVBoxLayout(log_gb)
        log_layout.setSpacing(4)
        log_layout.setContentsMargins(0, 0, 0, 0)

        clear_row = QHBoxLayout()
        clear_row.addStretch()
        clear_btn = QPushButton("🗑 Xóa log")
        clear_btn.setObjectName("smallBtn")
        clear_btn.setCursor(Qt.PointingHandCursor)
        clear_btn.clicked.connect(self._clear_log)
        clear_row.addWidget(clear_btn)
        log_layout.addLayout(clear_row)

        self.log_view = QTextEdit()
        self.log_view.setObjectName("logView")
        self.log_view.setReadOnly(True)
        log_layout.addWidget(self.log_view, 1)

        layout.addWidget(log_gb, 1)

        # ── Results group ──
        res_gb = QGroupBox("📁 File kết quả (Click đúp để mở)")
        res_layout = QVBoxLayout(res_gb)
        res_layout.setSpacing(8)
        res_layout.setContentsMargins(0, 0, 0, 0)

        self.result_list = QListWidget()
        self.result_list.setObjectName("resultList")
        self.result_list.setMaximumHeight(100)
        self.result_list.itemDoubleClicked.connect(self._on_open_result)
        res_layout.addWidget(self.result_list)

        res_btns = QHBoxLayout()
        res_btns.addStretch()
        open_folder_btn = QPushButton("📂 Mở thư mục")
        open_folder_btn.setObjectName("smallBtn")
        open_folder_btn.setCursor(Qt.PointingHandCursor)
        open_folder_btn.clicked.connect(self._open_output_dir)
        res_btns.addWidget(open_folder_btn)
        
        clear_res_btn = QPushButton("🗑 Xóa danh sách")
        clear_res_btn.setObjectName("smallBtn")
        clear_res_btn.setCursor(Qt.PointingHandCursor)
        clear_res_btn.clicked.connect(self._clear_results)
        res_btns.addWidget(clear_res_btn)
        
        res_layout.addLayout(res_btns)

        layout.addWidget(res_gb)
        return w

    # ═══════════════════════════════════════════════════════
    # TAB 4: TEST THỦ CÔNG
    # ═══════════════════════════════════════════════════════
    def _build_test_tab(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        w = QWidget()
        w.setObjectName("scrollContent")
        layout = QVBoxLayout(w)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        # ── Group 1: Tính toán từ Picking list ──
        gb1 = QGroupBox("📊 Tính toán từ Picking list có sẵn")
        gb1_layout = QVBoxLayout(gb1)
        gb1_layout.setSpacing(8)

        btn_row1 = QHBoxLayout()
        btn_pick = QPushButton("📂 Chọn file Picking list...")
        btn_pick.setObjectName("browseBtn")
        btn_pick.setCursor(Qt.PointingHandCursor)
        btn_pick.clicked.connect(self._test_select_picking)
        btn_row1.addWidget(btn_pick)

        btn_clear1 = QPushButton("🗑 Xóa danh sách")
        btn_clear1.setObjectName("smallBtn")
        btn_clear1.setCursor(Qt.PointingHandCursor)
        btn_clear1.clicked.connect(lambda: self._test_picking_list.clear())
        btn_row1.addWidget(btn_clear1)
        btn_row1.addStretch()
        gb1_layout.addLayout(btn_row1)

        self._test_picking_list = QListWidget()
        self._test_picking_list.setMaximumHeight(100)
        self._test_picking_list.setObjectName("resultList")
        gb1_layout.addWidget(self._test_picking_list)

        btn_calc = QPushButton("▶ Chạy tính toán")
        btn_calc.setObjectName("schedBtn")
        btn_calc.setCursor(Qt.PointingHandCursor)
        btn_calc.clicked.connect(self._test_run_calculator)
        gb1_layout.addWidget(btn_calc)

        layout.addWidget(gb1)

        # ── Group 2: In shipping label ──
        gb2 = QGroupBox("🖨️ In Shipping label có sẵn")
        gb2_layout = QVBoxLayout(gb2)
        gb2_layout.setSpacing(8)

        btn_row2 = QHBoxLayout()
        btn_ship = QPushButton("📂 Chọn file Shipping label...")
        btn_ship.setObjectName("browseBtn")
        btn_ship.setCursor(Qt.PointingHandCursor)
        btn_ship.clicked.connect(self._test_select_shipping)
        btn_row2.addWidget(btn_ship)

        btn_clear2 = QPushButton("🗑 Xóa danh sách")
        btn_clear2.setObjectName("smallBtn")
        btn_clear2.setCursor(Qt.PointingHandCursor)
        btn_clear2.clicked.connect(lambda: self._test_shipping_list.clear())
        btn_row2.addWidget(btn_clear2)
        btn_row2.addStretch()
        gb2_layout.addLayout(btn_row2)

        self._test_shipping_list = QListWidget()
        self._test_shipping_list.setMaximumHeight(100)
        self._test_shipping_list.setObjectName("resultList")
        gb2_layout.addWidget(self._test_shipping_list)

        btn_print = QPushButton("🖨️ In + Merge 2-up")
        btn_print.setObjectName("schedBtn")
        btn_print.setCursor(Qt.PointingHandCursor)
        btn_print.clicked.connect(self._test_print_shipping)
        gb2_layout.addWidget(btn_print)

        layout.addWidget(gb2)

        # ── Group 3: Tự động toàn bộ (tính toán + in) ──
        gb3 = QGroupBox("🚀 Tự động toàn bộ (chọn tất cả file đã tải về)")
        gb3_layout = QVBoxLayout(gb3)
        gb3_layout.setSpacing(8)

        btn_auto = QPushButton("📂 Chọn tất cả file của 1 lần tải...")
        btn_auto.setObjectName("browseBtn")
        btn_auto.setCursor(Qt.PointingHandCursor)
        btn_auto.clicked.connect(self._test_select_all)
        gb3_layout.addWidget(btn_auto)

        self._test_all_list = QListWidget()
        self._test_all_list.setMaximumHeight(100)
        self._test_all_list.setObjectName("resultList")
        gb3_layout.addWidget(self._test_all_list)

        btn_row3 = QHBoxLayout()
        btn_run_all = QPushButton("▶ Chạy toàn bộ (Tính toán + In)")
        btn_run_all.setObjectName("schedBtn")
        btn_run_all.setCursor(Qt.PointingHandCursor)
        btn_run_all.clicked.connect(self._test_run_all)
        btn_row3.addWidget(btn_run_all, 2)

        btn_clear3 = QPushButton("🗑 Xóa danh sách")
        btn_clear3.setObjectName("smallBtn")
        btn_clear3.setCursor(Qt.PointingHandCursor)
        btn_clear3.clicked.connect(lambda: self._test_all_list.clear())
        btn_row3.addWidget(btn_clear3)
        gb3_layout.addLayout(btn_row3)

        layout.addWidget(gb3)
        layout.addStretch()
        scroll.setWidget(w)
        return scroll

    def _test_select_picking(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Chọn file Picking list", "",
                                                 "PDF Files (*.pdf)")
        for f in files:
            self._test_picking_list.addItem(f)

    def _test_select_shipping(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Chọn file Shipping label", "",
                                                 "PDF Files (*.pdf)")
        for f in files:
            self._test_shipping_list.addItem(f)

    def _test_run_calculator(self):
        pdfs = [self._test_picking_list.item(i).text()
                for i in range(self._test_picking_list.count())]
        if not pdfs:
            QMessageBox.warning(self, "Cảnh báo", "Chọn ít nhất 1 file Picking list.")
            return
        master = self._master_real
        retail = self._retail_real
        if not master or not Path(master).exists():
            QMessageBox.critical(self, "Lỗi", "Chọn file Master (Combo) hợp lệ ở tab Tệp dữ liệu.")
            return
        out_dir = self.output_row.get_real_path() or str(BASE_DIR / "outputs")

        def _run():
            self._test_log.emit("info", "🧪 TEST: Bắt đầu tính toán...")
            try:
                results = run_calculator(pdfs, out_dir, master, retail,
                                         lambda m, t='': self._test_log.emit(t, m))
                for r in results:
                    self._test_log.emit("ok", f"  ✓ {r['rows']} SKU | Qty={r['tong_qty']} | Sold={r['tong_sold']} | Promo={r['tong_promo']}")
                    for key, lbl in [('pdf_report', '📄')]:
                        fp = r['files'].get(key)
                        if fp and Path(fp).exists():
                            self._add_result(fp, label=lbl)
                self._test_log.emit("bold_ok", "✅ Tính toán hoàn tất")
            except Exception as e:
                self._test_log.emit("err", f"✗ Lỗi: {e}")

        threading.Thread(target=_run, daemon=True).start()

    def _test_print_shipping(self):
        files = [self._test_shipping_list.item(i).text()
                 for i in range(self._test_shipping_list.count())]
        if not files:
            QMessageBox.warning(self, "Cảnh báo", "Chọn ít nhất 1 file Shipping label.")
            return
        printer = self.printer_combo.currentText()
        if not printer:
            QMessageBox.warning(self, "Cảnh báo", "Chọn máy in ở tab Cấu hình In.")
            return
        if not auto_print:
            QMessageBox.warning(self, "Cảnh báo", "Tick 'In tự động ra máy in' ở tab Cấu hình In.")
            return
        pdf_settings = self._build_pdf_settings()

        def _run():
            self._test_log.emit("info", f"🧪 TEST: In {len(files)} file...")
            for fp in files:
                try:
                    _print_file(fp, printer, pdf_settings=pdf_settings, batch_size=batch_size,
                                log_cb=lambda m, t='': self._test_log.emit(t, m))
                    self._test_log.emit("ok", f"  ✓ Đã gửi in: {Path(fp).name}")
                except Exception as e:
                    self._test_log.emit("err", f"  ✗ Lỗi in {Path(fp).name}: {e}")
            self._test_log.emit("bold_ok", "✅ In hoàn tất")

        threading.Thread(target=_run, daemon=True).start()

    def _test_select_all(self):
        """Chọn tất cả file PDF đã tải về (cả Picking list + Shipping label)."""
        files, _ = QFileDialog.getOpenFileNames(self, "Chọn tất cả file đã tải về", "",
                                                 "PDF Files (*.pdf)")
        for f in files:
            self._test_all_list.addItem(f)

    def _test_run_all(self):
        """Tự động: phân loại Picking list / Shipping label → tính toán → in."""
        all_files = [self._test_all_list.item(i).text()
                     for i in range(self._test_all_list.count())]
        if not all_files:
            QMessageBox.warning(self, "Cảnh báo", "Chọn ít nhất 1 file PDF.")
            return

        # Phân loại file khác
        other = [f for f in all_files if 'picking' not in Path(f).name.lower()
                 and 'shipping' not in Path(f).name.lower()
                 and 'vận chuyển' not in Path(f).name.lower()]

        master = self._master_real
        retail = self._retail_real
        out_dir = str(Path(all_files[0]).parent) if all_files else self.output_row.get_real_path() or str(BASE_DIR / "outputs")
        printer = self.printer_combo.currentText()
        auto_print = self.auto_print_cb.isChecked()
        pdf_settings = self._build_pdf_settings()
        batch_size = self.batch_size_spin.value()

        def _run():
            # Gom tất cả file theo carrier
            do_print = auto_print and printer
            if not do_print:
                self._test_log.emit("warn", "⚠ In bị tắt — tick 'In tự động ra máy in' ở tab Cấu hình In để in")
            carrier_map = {'JnT': 'J&T', 'GHN': 'GHN', 'VietNam': 'VietNam Post',
                           'Best': 'Best Express', 'Viettel': 'Viettel Post', 'JTC': 'J&T Cargo VN'}

            # Nhóm file theo carrier: {carrier: {picking: [...], shipping: [...]}}
            by_carrier = {}
            for f in all_files:
                fname = Path(f).name
                prefix = fname.split('_')[0]
                carrier = carrier_map.get(prefix, prefix)
                if carrier not in by_carrier:
                    by_carrier[carrier] = {'picking': [], 'shipping': []}
                if 'picking' in fname.lower():
                    by_carrier[carrier]['picking'].append(f)
                elif 'shipping' in fname.lower() or 'vận chuyển' in fname.lower():
                    by_carrier[carrier]['shipping'].append(f)

            # Xử lý từng carrier: shipping → tính toán → báo cáo
            for carrier, groups in by_carrier.items():
                self._test_log.emit("info", f"─── [{carrier}] ───")

                # 1. In shipping label trước
                if do_print:
                    for fp in groups['shipping']:
                        try:
                            _print_file(fp, printer, pdf_settings=pdf_settings, batch_size=batch_size,
                                        log_cb=lambda m, t='': self._test_log.emit(t, m))
                            self._test_log.emit("ok", f"  🖨️ {Path(fp).name}")
                        except Exception as e:
                            self._test_log.emit("err", f"  ✗ Lỗi in {Path(fp).name}: {e}")

                # 2. Tính toán (luôn chạy dù có in hay không)
                results = []
                if groups['picking']:
                    self._test_log.emit("info", f"  📊 {len(groups['picking'])} Picking list → tính toán...")
                    try:
                        results = run_calculator(groups['picking'], out_dir, master, retail,
                                                 lambda m, t='': self._test_log.emit(t, m),
                                                 carrier=carrier)
                        for r in results:
                            self._test_log.emit("ok", f"  ✓ {r['rows']} SKU | Qty={r['tong_qty']}")
                            fp = r['files'].get('pdf_report')
                            if fp and Path(fp).exists():
                                self._add_result(fp)
                                if do_print:
                                    try:
                                        _print_file(fp, printer, pdf_settings=pdf_settings, batch_size=batch_size,
                                                    log_cb=lambda m, t='': self._test_log.emit(t, m))
                                        self._test_log.emit("ok", f"  🖨️ Báo cáo: {Path(fp).name}")
                                    except Exception as e:
                                        self._test_log.emit("err", f"  ✗ Lỗi in báo cáo: {e}")
                    except Exception as e:
                        self._test_log.emit("err", f"  ✗ Lỗi tính toán [{carrier}]: {e}")

            if other:
                self._test_log.emit("warn", f"⚠ {len(other)} file không rõ loại, bỏ qua")

            self._test_log.emit("bold_ok", "✅ Hoàn tất toàn bộ")

        threading.Thread(target=_run, daemon=True).start()

    # ═══════════════════════════════════════════════════════
    # BOTTOM BAR
    # ═══════════════════════════════════════════════════════
    def _build_bottom_bar(self):
        bar = QWidget()
        bar.setObjectName("bottomBar")
        layout = QVBoxLayout(bar)
        layout.setContentsMargins(24, 16, 24, 16)
        layout.setSpacing(12)

        self.status_label = QLabel("✅ Hệ thống sẵn sàng")
        self.status_label.setObjectName("statusLabel")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("color: #059669;")

        layout.addWidget(self.status_label)

        btn_row = QHBoxLayout()
        btn_row.setSpacing(12)

        self.sched_btn = QPushButton("▶ CHẠY")
        self.sched_btn.setObjectName("schedBtn")
        self.sched_btn.setCursor(Qt.PointingHandCursor)
        self.sched_btn.clicked.connect(self._on_run_schedule)
        btn_row.addWidget(self.sched_btn, 2)

        self.stop_btn = QPushButton("⏹ DỪNG")
        self.stop_btn.setObjectName("stopBtn")
        self.stop_btn.setCursor(Qt.PointingHandCursor)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self._on_stop)
        btn_row.addWidget(self.stop_btn, 1)

        layout.addLayout(btn_row)
        return bar

    # ═══════════════════════════════════════════════════════
    # STATUS UPDATES
    # ═══════════════════════════════════════════════════════
    def _update_cookie_status(self):
        self.cookie_row.update_cookie_status()

    def _update_master_status(self):
        self.master_row.update_excel_status()

    def _update_retail_status(self):
        self.retail_row.update_excel_status()

    def _on_cookie_changed(self, path: str):
        self._cookie_real = path
        self._update_cookie_status()

    def _on_master_changed(self, path: str):
        self._master_real = path
        self._update_master_status()
        self._load_preview_data()

    def _on_retail_changed(self, path: str):
        self._retail_real = path
        self._update_retail_status()
        self._load_preview_data()

    def _on_output_dir_changed(self, path: str):
        os.makedirs(path, exist_ok=True)

    # ═══════════════════════════════════════════════════════
    # PREVIEW TABLE LOGIC WITH SEARCH
    # ═══════════════════════════════════════════════════════
    def _load_preview_data(self):
        mp = self._master_real
        if mp and Path(mp).exists():
            self._populate_table(self.master_table, mp, "master")
        rp = self._retail_real
        if rp and Path(rp).exists():
            self._populate_table(self.retail_table, rp, "retail")

    def _populate_table(self, table: QTableWidget, file_path: str, table_type: str):
        try:
            from openpyxl import load_workbook
            wb = load_workbook(file_path, data_only=True, read_only=True)
            ws = wb.active

            headers = [str(cell.value) if cell.value else f"Col{i}"
                       for i, cell in enumerate(ws[1], 1)]

            if table_type == "master":
                cols = headers[:6] if len(headers) >= 6 else headers
            else:
                cols = headers[:4] if len(headers) >= 4 else headers

            # Block signals để cellChanged không fire khi đang load dữ liệu
            table.blockSignals(True)

            table.setColumnCount(len(cols))
            table.setHorizontalHeaderLabels(cols)

            rows_data = []
            excel_rows_list = []
            max_rows = 5001  # Nâng hạn mức load lên 5000 dòng
            for i, row in enumerate(ws.iter_rows(min_row=2, max_row=max_rows, values_only=True)):
                if any(cell is not None for cell in row):
                    rows_data.append(row)
                    excel_rows_list.append(i + 2)  # Hàng thực trong Excel (1-indexed)

            table.setRowCount(len(rows_data))
            for ri, row in enumerate(rows_data):
                for ci in range(len(cols)):
                    val = str(row[ci]) if ci < len(row) and row[ci] is not None else ""
                    item = QTableWidgetItem(val)
                    table.setItem(ri, ci, item)

            # Lưu metadata để cellChanged handler biết ghi vào đâu
            table._file_path = file_path
            table._excel_rows = excel_rows_list

            wb.close()
            table.blockSignals(False)
            table.resizeColumnsToContents()
            table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        except Exception:
            pass

    def _on_table_cell_changed(self, row: int, col: int):
        """Khi người dùng sửa một ô → tự động ghi vào file Excel gốc."""
        table = self.sender()
        file_path = getattr(table, '_file_path', None)
        excel_rows = getattr(table, '_excel_rows', [])
        if not file_path or row >= len(excel_rows):
            return
        excel_row = excel_rows[row]  # Hàng thực trong Excel (1-indexed)
        item = table.item(row, col)
        if item is None:
            return
        raw_text = item.text().strip()
        # Tự động đoán kiểu dữ liệu: int → float → string
        if raw_text == "":
            new_value = None
        else:
            try:
                new_value = int(raw_text)
            except ValueError:
                try:
                    new_value = float(raw_text)
                except ValueError:
                    new_value = raw_text
        try:
            from openpyxl import load_workbook
            wb = load_workbook(file_path)
            ws = wb.active
            ws.cell(row=excel_row, column=col + 1, value=new_value)
            wb.save(file_path)
            wb.close()
        except Exception:
            pass  # File đang mở bởi ứng dụng khác → bỏ qua

    def _on_preview_tab_changed(self, idx: int):
        is_master = (idx == 0)
        self.master_table.setVisible(is_master)
        self.retail_table.setVisible(not is_master)
        
        for i, btn in enumerate([self.preview_master_btn, self.preview_retail_btn]):
            btn.setProperty("active", i == idx)
            btn.style().unpolish(btn)
            btn.style().polish(btn)
            
        for table in [self.master_table, self.retail_table]:
            for row in range(table.rowCount()):
                table.setRowHidden(row, False)
        
        self._filter_preview_table(self.search_input.text())

    def _filter_preview_table(self, text: str):
        text_lower = text.lower().strip()
        active_table = self.master_table if self.preview_btn_group.checkedId() == 0 else self.retail_table
        
        for row in range(active_table.rowCount()):
            if not text_lower:
                active_table.setRowHidden(row, False)
                continue
            match = False
            for col in range(active_table.columnCount()):
                item = active_table.item(row, col)
                if item and text_lower in item.text().lower():
                    match = True
                    break
            active_table.setRowHidden(row, not match)

    # ═══════════════════════════════════════════════════════
    # BUTTON STATE MANAGEMENT
    # ═══════════════════════════════════════════════════════
    def _set_buttons(self, state: str):
        if state in ("running", "scheduled"):
            self.sched_btn.setEnabled(False)
            self.stop_btn.setEnabled(True)
        else:
            self.sched_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)

    # ═══════════════════════════════════════════════════════
    # CONFIG COLLECTOR
    # ═══════════════════════════════════════════════════════
    def _collect_config(self) -> dict:
        carriers_to_process = []
        carrier_keys = ["jt", "ghn", "vnp", "best", "viettel", "jtc"]
        carrier_names = ["J&T", "GHN", "VietNam Post", "Best Express", "Viettel Post", "J&T Cargo VN"]
        for key, name in zip(carrier_keys, carrier_names):
            if self.carrier_checkboxes[key].isChecked():
                val = self.carrier_spinboxes[key].value()
                carriers_to_process.append((name, val))
            # unchecked = bỏ qua hãng này hoàn toàn

        return {
            "cookie": self._cookie_real,
            "output_dir": self.output_row.get_real_path() or str(BASE_DIR / "outputs"),
            "master": self._master_real,
            "retail": self._retail_real,
            "carriers": carriers_to_process,
            "auto_print": self.auto_print_cb.isChecked(),
            "printer": self.printer_combo.currentText(),
            "test_mode": self.test_mode_cb.isChecked(),
            "batch_size": self.batch_size_spin.value(),
            "pdf_settings": self._build_pdf_settings(),
        }

    def _build_pdf_settings(self) -> str:
        parts = ["paper=A4",
                 f"duplex={self.duplex_combo.currentText()}"]
        return ",".join(parts)

    # ═══════════════════════════════════════════════════════
    # PRINT HELPERS
    # ═══════════════════════════════════════════════════════
    def _refresh_printers(self):
        printers = _get_printers()
        self.printer_combo.clear()
        self.printer_combo.addItems(printers)
        if printers:
            self.printer_combo.setCurrentIndex(0)

    def _on_sumatra_path_changed(self, text: str):
        global _custom_sumatra_path
        _custom_sumatra_path = text.strip()

    def _browse_sumatra_path(self):
        path, _ = QFileDialog.getOpenFileName(self, "Chọn SumatraPDF.exe", "",
                                              "SumatraPDF (SumatraPDF.exe);;All Files (*.*)")
        if path:
            self.sumatra_path_edit.setText(path)

    # ═══════════════════════════════════════════════════════
    # LOG (colored HTML via QTextEdit)
    # ═══════════════════════════════════════════════════════
    def _log_html(self, tag: str, msg: str):
        ts = datetime.now().strftime("%H:%M:%S")
        color = self.TAG_COLORS.get(tag, "#E2E8F0")
        weight = "font-weight: bold;" if "bold" in tag else ""
        html = (
            f"<span style='color:{color};'>{ts}  </span>"
            f"<span style='color:{color};{weight}'>{msg}</span><br>"
        )
        cursor = self.log_view.textCursor()
        cursor.movePosition(QTextCursor.End)
        cursor.insertHtml(html)
        # Giới hạn ~5000 dòng: xóa nửa đầu thay vì clear toàn bộ
        doc = self.log_view.document()
        if doc.blockCount() > 5000:
            remove_count = doc.blockCount() - 2500
            cursor = QTextCursor(doc.begin())
            for _ in range(remove_count):
                cursor.movePosition(QTextCursor.NextBlock, QTextCursor.KeepAnchor)
            cursor.removeSelectedText()
            # Thông báo đã trim log
            cursor.movePosition(QTextCursor.Start)
            cursor.insertHtml(
                f"<span style='color:#64748B;'>[... đã xóa {remove_count} dòng cũ — giữ 2500 dòng gần nhất ...]</span><br>"
            )
            self.log_view.ensureCursorVisible()

    @Slot(str, str)
    def _on_log_message(self, tag: str, msg: str):
        self._log_html(tag, msg)

    @Slot(str, str)
    def _on_state_changed(self, step: str, msg: str):
        colors = {
            "idle": "#059669", "running": "#D97706",
            "waiting": "#2563EB", "done": "#059669", "error": "#DC2626",
        }
        color = colors.get(step, "#1E293B")
        self.status_label.setText(msg)
        self.status_label.setStyleSheet(f"color: {color}; font-weight: bold; font-size: 14px;")

    @Slot(object)
    def _on_job_completed(self, result: dict):
        self.running = False
        # Ghi nhận thời điểm hoàn thành cho scheduler interval (tính từ lúc KẾT THÚC)
        if self.scheduler_active:
            self._sched_last_run = datetime.now()
        else:
            self._set_buttons("idle")
        if result:
            for p in result.get('pdf_paths', []):
                if Path(p).exists():
                    self._add_result(p)
            for r in result.get('results', []):
                for key, lbl in [('excel', '📊'), ('pdf', '📕'), ('pdf_grouped', '📋')]:
                    fp = r['files'].get(key)
                    if fp and Path(fp).exists():
                        self._add_result(fp, label=lbl)

    def _clear_log(self):
        self.log_view.clear()

    # ═══════════════════════════════════════════════════════
    # RESULTS
    # ═══════════════════════════════════════════════════════
    def _add_result(self, file_path: str, label: str = ''):
        """Thêm file vào danh sách kết quả. file_path là đường dẫn đầy đủ."""
        display = f'{label} {Path(file_path).name}' if label else f'📄 {Path(file_path).name}'
        self.result_list.addItem(display)
        self.result_list.item(self.result_list.count() - 1).setData(Qt.UserRole, file_path)
        self.result_files.append(file_path)

    def _clear_results(self):
        self.result_list.clear()
        self.result_files.clear()

    def _on_open_result(self, item: QListWidgetItem):
        fp = item.data(Qt.UserRole)
        if fp and Path(fp).exists():
            os.startfile(fp)

    def _open_output_dir(self):
        base = self.output_row.get_real_path() or str(BASE_DIR / "outputs")
        d = Path(base) / datetime.now().strftime("%Y-%m-%d")
        if d.exists():
            os.startfile(str(d))
        elif Path(base).exists():
            os.startfile(base)

    # ═══════════════════════════════════════════════════════
    # ACTION CALLBACKS
    # ═══════════════════════════════════════════════════════
    def _on_run_now(self):
        if self.running:
            return
        if not Path(self._cookie_real).exists():
            QMessageBox.critical(self, "Lỗi", "Chọn file cookie JSON hợp lệ.")
            return
        config = self._collect_config()
        if not config['carriers']:
            QMessageBox.warning(self, "Cảnh báo", "Vui lòng chọn ít nhất 1 hãng vận chuyển để in.")
            return
        self.running = True
        self._set_buttons("running")
        self._clear_results()
        self._clear_log()
        self._log_html("bold_ok", "▶ Bắt đầu...")
        self.status_label.setText("⏳ Đang xử lý...")
        self.status_label.setStyleSheet("color: #D97706; font-weight: bold; font-size: 14px;")
        self.trigger_job.emit(config)

    def _on_run_schedule(self):
        if not Path(self._cookie_real).exists():
            QMessageBox.critical(self, "Lỗi", "Chọn file cookie JSON hợp lệ.")
            return
        mode = self._sched_mode
        if mode == "once":
            self._on_run_now()
            return
        if mode == "daily":
            times = [t.strip() for t in self.daily_times_edit.text().split(",") if t.strip()]
            if not times:
                QMessageBox.critical(self, "Lỗi", "Nhập ít nhất 1 giờ (VD: 08:00).")
                return
            self._sched_daily_times = []
            for t in times:
                try:
                    h, m = t.strip().split(":")
                    self._sched_daily_times.append((int(h), int(m)))
                except ValueError:
                    QMessageBox.critical(self, "Lỗi", f"Giờ không hợp lệ: {t}")
                    return
        elif mode == "interval":
            self._sched_interval_hours = self.interval_spin.value()

        self._clear_log()
        self._clear_results()
        self._log_html("info", f"⏰ Hẹn giờ: {mode}")
        if mode == "interval":
            self._log_html("info", f"   Chạy mỗi {self._sched_interval_hours} giờ")
        elif mode == "daily":
            self._log_html("info", f"   Chạy lúc {self.daily_times_edit.text()}")

        self._sched_mode = mode
        self.scheduler_active = True
        self._sched_last_run = None
        self._set_buttons("scheduled")
        self._sched_timer.start()
        self._update_countdown()

    def _on_stop(self):
        self.scheduler_active = False
        self._sched_timer.stop()
        self.running = False
        # Dùng invokeMethod để gửi lệnh stop qua event queue của worker thread (đúng chuẩn Qt)
        # threading.Event.set() là thread-safe nên nếu invokeMethod thất bại, direct call vẫn an toàn
        if not QMetaObject.invokeMethod(self._worker, "stop_job", Qt.QueuedConnection):
            self._worker.stop_job()  # fallback an toàn vì chỉ set threading.Event
        self._set_buttons("idle")
        self.status_label.setText("⏹ Đã dừng")
        self.status_label.setStyleSheet("color: #DC2626; font-weight: bold; font-size: 14px;")
        self._log_html("warn", "⏹ Đã dừng hệ thống")

    def _on_schedule_mode_changed(self, btn: QRadioButton):
        mode = btn.property("mode")
        self._sched_mode = mode
        self.interval_panel.setVisible(mode == "interval")
        self.daily_panel.setVisible(mode == "daily")

    # ═══════════════════════════════════════════════════════
    # SCHEDULER (QTimer-based, main thread)
    # ═══════════════════════════════════════════════════════
    def _check_schedule(self):
        if not self.scheduler_active:
            return
        now = datetime.now()
        if self._sched_mode == "interval":
            # Lần đầu chạy ngay, các lần sau cách nhau N giờ tính từ lúc hoàn thành
            if self._sched_last_run is None:
                self._sched_last_run = now
                self._execute_scheduled_job()
                return
            next_run = self._sched_last_run + timedelta(hours=self._sched_interval_hours)
            self._sched_next_run = next_run
            if now >= next_run:
                self._execute_scheduled_job()
            else:
                self._update_countdown()
        elif self._sched_mode == "daily":
            candidates = []
            for h, m in self._sched_daily_times:
                rt = now.replace(hour=h, minute=m, second=0, microsecond=0)
                if rt < now:
                    rt += timedelta(days=1)
                candidates.append(rt)
            if candidates:
                next_run = min(candidates)
                self._sched_next_run = next_run
                diff = (next_run - now).total_seconds()
                if diff <= 1:
                    if not self._sched_last_run or (now - self._sched_last_run).total_seconds() > 60:
                        self._execute_scheduled_job()
                else:
                    self._update_countdown()

    def _execute_scheduled_job(self):
        if self.running:
            self._log_html("dim", "⏭ Bỏ qua chu kỳ — job trước vẫn đang chạy")
            return
        config = self._collect_config()
        if not config['carriers']:
            self._log_html("warn", "⏭ Bỏ qua chu kỳ — không có hãng nào được chọn")
            return
        self.running = True
        self._clear_results()
        self.status_label.setText("🔄 Đang chạy tác vụ tự động...")
        self.status_label.setStyleSheet("color: #D97706; font-weight: bold; font-size: 14px;")
        config = self._collect_config()
        self.trigger_job.emit(config)

    def _update_countdown(self):
        if self._sched_next_run:
            remaining = self._sched_next_run - datetime.now()
            secs = max(0, int(remaining.total_seconds()))
            h, m = secs // 3600, (secs % 3600) // 60
            self.status_label.setText(f"⏳ Chạy tiếp sau {h}h{m:02d}")
            self.status_label.setStyleSheet("color: #2563EB; font-weight: bold; font-size: 14px;")

    # ═══════════════════════════════════════════════════════
    # CLOSE EVENT
    # ═══════════════════════════════════════════════════════
    def closeEvent(self, event):
        self.scheduler_active = False
        self._sched_timer.stop()
        self.running = False
        # shutdown() chứa browser.close() + playwright.stop() là process-level, không phụ thuộc thread
        # Gọi trực tiếp là an toàn; invokeMethod + BlockingQueuedConnection có thể deadlock nếu worker đang bận
        self._worker.shutdown()
        self._worker_thread.quit()
        self._worker_thread.wait(5000)
        event.accept()


# ============================================================
# ENTRY POINT
# ============================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = App()
    window.show()
    sys.exit(app.exec())