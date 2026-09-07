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

# ── Ép stdout/stderr dùng UTF-8 — tránh crash khi print() gặp emoji
# (⚠, 📊...) trên console đang dùng bảng mã cp1252/cp1258 thay vì UTF-8.
# sys.stdout có thể là None khi build windowed (console=False) nên phải guard.
for _stream in (sys.stdout, sys.stderr):
    if _stream is not None and hasattr(_stream, 'reconfigure'):
        try: _stream.reconfigure(encoding='utf-8', errors='replace')
        except Exception: pass

# ═══════════════════════════════════════════════════════════
# PySide6 imports
# ═══════════════════════════════════════════════════════════
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QLineEdit, QPushButton, QGroupBox, QCheckBox, QRadioButton,
    QButtonGroup, QSpinBox, QComboBox, QTextEdit, QListWidget,
    QListWidgetItem, QScrollArea, QStackedWidget, QTableWidget, QTabWidget,
    QTableWidgetItem, QHeaderView, QFileDialog, QMessageBox,
    QFrame, QApplication,
)
from PySide6.QtCore import (
    Qt, Signal, Slot, QThread, QTimer, QObject, QMetaObject,
)
from PySide6.QtGui import (
    QFont, QTextCursor,
)

# ═══════════════════════════════════════════════════════════
# Paths & config
# ═══════════════════════════════════════════════════════════
BASE_DIR = Path(__file__).parent
BILL_DIR = BASE_DIR / 'bill_calculate'
UPLOAD_DIR = BILL_DIR / 'uploads'

# ═══════════════════════════════════════════════════════════
# Frozen / source mode — detect paths
# ═══════════════════════════════════════════════════════════
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).parent
    BILL_DIR = Path(sys._MEIPASS) / 'bill_calculate'
    UPLOAD_DIR = BILL_DIR / 'uploads'
    DEFAULT_COOKIE = Path(sys._MEIPASS) / 'seller-vn.tiktok.com_25-06-2026.json'
    MASTER_DEFAULT = Path(sys._MEIPASS) / 'mã combo.xlsx'
    RETAIL_DEFAULT = Path(sys._MEIPASS) / 'sp bán lẻ.xlsx'
    TEMPLATE_DEFAULT = Path(sys._MEIPASS) / 'Bảng thống kê hàng.xlsx'
else:
    DEFAULT_COOKIE = BASE_DIR / 'seller-vn.tiktok.com_25-06-2026.json'
    MASTER_DEFAULT = BASE_DIR / 'mã combo.xlsx'
    RETAIL_DEFAULT = BASE_DIR / 'sp bán lẻ.xlsx'
    TEMPLATE_DEFAULT = BASE_DIR / 'Bảng thống kê hàng.xlsx'

sys.path.insert(0, str(BILL_DIR))

TARGET_URL = 'https://seller-vn.tiktok.com'
ORDERS_URL = 'https://seller-vn.tiktok.com/order?order_status%5B%5D=1&selected_sort=11&tab=to_ship&page_size=50'
CHROME_USER_DATA_DIR = BASE_DIR / '.chrome_profile'
os.makedirs(CHROME_USER_DATA_DIR, exist_ok=True)
CARRIER_URLS = {
    'GHN':           ORDERS_URL + '&shipping_provider_id%5B%5D=7252807945006614278',
    'J&T':           ORDERS_URL + '&shipping_provider_id%5B%5D=6841743441349706241',
    'VietNam Post':  ORDERS_URL + '&shipping_provider_id%5B%5D=7062208235196909313',
    'Best Express':  ORDERS_URL + '&shipping_provider_id%5B%5D=7099655686241388293',
    'Viettel Post':  ORDERS_URL + '&shipping_provider_id%5B%5D=7155825439565416197',
    'J&T Cargo VN':  ORDERS_URL + '&shipping_provider_id%5B%5D=7581675938962736917',
}
BATCH_SIZE = 50
API_IDS_URL = 'http://88.2.0.55:7016/api/ids/receive'  # Endpoint nhận Order ID
os.makedirs(UPLOAD_DIR, exist_ok=True)


# ============================================================
# AUTOMATION
# ============================================================
def _detect_captcha(page, log_cb, state_cb, stop_event, output_dir=''):
    """Kiểm tra xem TikTok có hiện CAPTCHA không. Nếu có → dừng chờ user giải."""
    captcha_selectors = [
        # TikTok slider CAPTCHA
        'iframe[src*="captcha"]',
        'iframe[src*="verify"]',
        'div[class*="captcha"]',
        'div[class*="verify"]',
        'div[class*="slider"]',
        # Text-based detection
        'text=Kéo thanh trượt',
        'text=Kéo để xác',
        'text=Trượt để xác',
        'text=Slide to verify',
        'text=Please verify',
        'text=Xác minh',
        # Common CAPTCHA container IDs
        '#captcha',
        '#captcha-container',
        '.captcha_verify',
        '[data-testid="captcha"]',
    ]
    for selector in captcha_selectors:
        try:
            el = page.locator(selector).first
            if el.count() > 0 and el.is_visible(timeout=1000):
                # Có CAPTCHA!
                log_cb('🛑 PHÁT HIỆN CAPTCHA! Vui lòng kéo hình xác minh trên Chrome...', 'err')
                state_cb('captcha', '⏳ Đợi bạn giải CAPTCHA...')
                # Chụp màn hình
                try:
                    ss_dir = output_dir if output_dir else '.'
                    ss = str(Path(ss_dir) / f'captcha_{datetime.now().strftime("%m-%d_%H-%M-%S")}.png')
                    page.screenshot(path=ss)
                    log_cb(f'  📸 Screenshot: {ss}', 'info')
                except: pass
                # Đợi user giải CAPTCHA (polling mỗi 2s, tối đa 5 phút)
                import time as _t
                for _ in range(150):  # 150 × 2s = 5 phút
                    if stop_event and stop_event.is_set():
                        return
                    _t.sleep(2)
                    # Kiểm tra CAPTCHA đã biến mất chưa
                    try:
                        if el.count() == 0 or not el.is_visible(timeout=500):
                            log_cb('✅ CAPTCHA đã được giải — đợi 5s để trang load lại...', 'ok')
                            state_cb('running', 'Đang đợi trang load lại...')
                            page.wait_for_timeout(5000)
                            state_cb('running', 'Đang tiếp tục...')
                            return
                    except:
                        log_cb('✅ CAPTCHA đã được giải — đợi 5s để trang load lại...', 'ok')
                        state_cb('running', 'Đang đợi trang load lại...')
                        page.wait_for_timeout(5000)
                        state_cb('running', 'Đang tiếp tục...')
                        return
                log_cb('⚠ Hết thời gian chờ CAPTCHA — thử tiếp...', 'warn')
                return
        except Exception:
            continue

def run_automation(cookie_path, output_dir, max_orders, log_cb, state_cb, stop_event=None,
                   existing_playwright=None, existing_browser=None, carrier=None, test_mode=False,
                   exclude_pre_orders=True):
    with open(cookie_path, 'r', encoding='utf-8') as f:
        cd = json.load(f)
    cookies_list = cd.get('cookies', cd if isinstance(cd, list) else [])
    pdf_files = []
    total_printed = 0
    target = max_orders if max_orders > 0 else 10**9
    use_select_all = (max_orders == 0)
    batch_num = 0
    select_all_batches = 0  # Đếm số batch khi dùng "Chọn tất cả" — chống loop vô hạn
    carrier_label = f' [{carrier}]' if carrier else ''
    orders_url = CARRIER_URLS.get(carrier, ORDERS_URL)
    if exclude_pre_orders:
        orders_url += '&order_exclusion%5B%5D=1'  # Loại trừ đơn bán trước
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
            _detect_captcha(page, log_cb, state_cb, stop_event, output_dir)
            browser_ok = True
        except Exception as e:
            log_cb(f'⚠ Không dùng lại được browser cũ ({e}) — tạo mới...', 'warn')
            # QUAN TRỌNG: chỉ close browser, KHÔNG stop playwright
            # stop() gọi vào native code dễ gây crash nếu driver đang dở việc
            try:
                existing_browser.close()
            except Exception:
                pass
            # Đợi browser process thoát hẳn trước khi tạo mới
            import time as _time
            _time.sleep(0.5)
            existing_browser = None
            existing_playwright = None

    if not browser_ok:
        # ── Stop playwright cũ TRƯỚC KHI tạo mới ──
        # Nếu không stop, sync_playwright().start() sẽ phát hiện event loop cũ
        # còn chạy và ném lỗi: "using Playwright Sync API inside the asyncio loop"
        if existing_playwright is not None:
            try:
                existing_playwright.stop()
            except Exception:
                pass
            existing_playwright = None
            import time as _time_stop
            _time_stop.sleep(0.3)  # Đợi event loop cũ thoát hẳn

        try:
            from playwright.sync_api import sync_playwright  # Lazy import — chỉ load khi chạy automation
            playwright = sync_playwright().start()
            # Các cờ an toàn để giảm tải nền (extension/sync/telemetry...) — KHÔNG đụng tới
            # GPU/renderer vì máy đang chạy GPU thật (Intel UHD qua D3D11), tắt GPU sẽ làm
            # mất WebGL/Canvas và khiến cuộn/kéo giật NẶNG hơn, đã kiểm chứng thực tế.
            chrome_args = [
                '--disable-blink-features=AutomationControlled',
                '--disable-extensions',
                '--disable-background-networking',
                '--disable-background-timer-throttling',
                '--disable-backgrounding-occluded-windows',
                '--disable-breakpad',
                '--disable-client-side-phishing-detection',
                '--disable-default-apps',
                '--disable-hang-monitor',
                '--disable-ipc-flooding-protection',
                '--disable-popup-blocking',
                '--disable-prompt-on-repost',
                '--disable-renderer-backgrounding',
                '--disable-sync',
                '--metrics-recording-only',
                '--no-first-run',
                '--password-store=basic',
                '--use-mock-keychain',
                '--mute-audio',
                '--disable-infobars',
                '--no-default-browser-check',
                '--no-pings',
                '--safebrowsing-disable-auto-update',
                '--disable-search-engine-choice-screen',
            ]
            log_cb('🌐 Dùng Google Chrome có sẵn trên máy (profile bền vững — cache giữ lại giữa các lần chạy)', 'info')
            # launch_persistent_context: dùng 1 thư mục profile cố định thay vì
            # tạo profile trắng mỗi lần → giữ cache DNS/JS/CSS/hình ảnh, load nhanh
            # hơn từ lần chạy thứ 2. context.browser vẫn dùng is_connected()/contexts
            # bình thường nên không phá logic tái sử dụng browser giữa các carrier.
            context = playwright.chromium.launch_persistent_context(
                str(CHROME_USER_DATA_DIR),
                headless=False, channel='chrome', args=chrome_args,
                viewport={'width': 1366, 'height': 768},
                user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/131.0.0.0 Safari/537.36',
                accept_downloads=True)
            browser = context.browser
        except Exception as e:
            log_cb(f'✗ Không thể khởi động browser: {e}', 'err')
            raise RuntimeError(f'Không thể khởi động Chromium: {e}') from e
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
        # launch_persistent_context tự mở sẵn 1 tab trắng — dùng lại tab đó thay vì
        # mở thêm tab mới rồi để tab trắng chạy nền vô ích.
        page = context.pages[0] if context.pages else context.new_page()
        for p in list(context.pages):
            if p != page:
                try:
                    if not p.is_closed(): p.close()
                except Exception: pass
        page.goto(TARGET_URL, wait_until='domcontentloaded', timeout=30000)
        page.wait_for_timeout(2000)
        _detect_captcha(page, log_cb, state_cb, stop_event, output_dir)

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
            _detect_captcha(page, log_cb, state_cb, stop_event, output_dir)

            total_avail = page.evaluate("() => document.querySelectorAll('td.col-checkbox label.p-checkbox').length")
            if total_avail == 0:
                # Kiểm tra có phải do chưa đăng nhập hay thực sự hết đơn
                try:
                    is_logged_out = page.evaluate("""() => {
                        // Check API-based session: page title contains login
                        if (document.title && document.title.toLowerCase().includes('login')) {
                            return true;
                        }
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
                page.wait_for_timeout(3500)

                # ── Đếm số đơn đã được chọn sau khi click header checkbox ──
                # Nếu trang không đầy (số đơn ít hơn page_size) → tất cả đơn
                # đã hiển thị trên 1 trang, không cần bấm "Chọn tất cả".
                # Bấm "Chọn tất cả" khi trang không đầy sẽ bị TOGGLE OFF → mất hết checkbox!
                header_checked = page.evaluate("""() => {
                    // Đếm row đã chọn — KHÔNG đếm header checkbox (trong thead/th)
                    // Dùng tr.p-highlight trước vì header row không có class này
                    let n = document.querySelectorAll('tr.p-highlight, tr.p-selectable-row.p-highlight').length;
                    if (n > 0) return n;
                    // Chỉ đếm aria-checked trong tbody (loại trừ header)
                    n = document.querySelectorAll('tbody [aria-checked="true"], tr[aria-checked="true"]').length;
                    if (n > 0) return n;
                    // Chỉ đếm p-checkbox-checked trong td (loại trừ th)
                    n = document.querySelectorAll('td .p-checkbox-checked').length;
                    if (n > 0) return n;
                    n = document.querySelectorAll('td.col-checkbox svg').length;
                    if (n > 0) return n;
                    // Chỉ đếm checkbox đã check trong tbody
                    return document.querySelectorAll('tbody input[type="checkbox"]:checked').length;
                }""")
                log_cb(f'  📋 Header checkbox đã chọn {header_checked}/{total_avail} đơn', 'info')

                # ── Phân biệt Case 1 vs Case 2 dựa vào sự TỒN TẠI của nút ──
                # TikTok chỉ có 2 loại nút: "Chọn X đơn hàng đầu tiên" hoặc "Chọn tất cả X đơn hàng"
                # Cả 2 đều chứa "Chọn" + "đơn hàng" — dùng cả 2 từ để tránh nhầm nút khác
                # ("Xuất đơn hàng", "Lọc đơn hàng"... có "đơn hàng" nhưng không có "Chọn")

                # ── Tìm button chứa "đơn hàng", sau đó lọc thêm "chọn" ──
                select_all_btn = None
                try:
                    candidates = page.locator('button:has-text("đơn hàng")').all()
                    for b in candidates:
                        try:
                            txt = b.inner_text().strip().lower()
                            if 'chọn' in txt and b.is_visible():
                                select_all_btn = b
                                break
                        except Exception:
                            pass
                except Exception:
                    pass

                if select_all_btn:
                    # ✅ Case 2: Tìm thấy nút "Chọn tất cả" → có >20 đơn
                    btn_text = select_all_btn.inner_text().strip()[:60]
                    select_all_btn.click(timeout=5000)
                    log_cb(f'  ✅ Đã bấm "{btn_text}"', 'ok')
                    page.wait_for_timeout(3000)
                    checked = page.evaluate("""() => {
                        let n = document.querySelectorAll('tr.p-highlight, tr.p-selectable-row.p-highlight').length;
                        if (n > 0) return n;
                        n = document.querySelectorAll('tbody [aria-checked="true"], tr[aria-checked="true"]').length;
                        if (n > 0) return n;
                        n = document.querySelectorAll('td .p-checkbox-checked').length;
                        if (n > 0) return n;
                        n = document.querySelectorAll('td.col-checkbox svg').length;
                        if (n > 0) return n;
                        const bar = document.querySelector('[class*="selected"], [class*="Selected"], [class*="count"]');
                        if (bar) {
                            const m = bar.textContent.match(/(\\d+)\\s*đơn/);
                            if (m) return parseInt(m[1]);
                        }
                        return document.querySelectorAll('tbody input[type="checkbox"]:checked').length;
                    }""")
                    log_cb(f'  ✓ Đã chọn {checked} đơn hàng', 'ok')
                    if checked < header_checked:
                        log_cb(f'  ⚠ "Chọn tất cả" đã toggle off ({header_checked}→{checked}) — chọn lại bằng header checkbox...', 'warn')
                        for hdr_sel in ['th .p-checkbox', 'th.col-checkbox .p-checkbox', 'th input[type="checkbox"]']:
                            try:
                                hdr = page.locator(hdr_sel).first
                                if hdr.count() > 0 and hdr.is_visible(timeout=1000):
                                    hdr.click()
                                    page.wait_for_timeout(2000)
                                    break
                            except: pass
                        checked = header_checked
                        log_cb(f'  ✓ Đã chọn lại {checked} đơn hàng', 'ok')

                    # ── Phân biệt 2 loại nút ──
                    # "Chọn X đơn hàng ĐẦU TIÊN" → còn đơn, chạy tiếp
                    # "Chọn TẤT CẢ X đơn hàng" → hết đơn, dừng
                    btn_text_lower = btn_text.lower()
                    if 'tất cả' in btn_text_lower:
                        force_stop = True
                        log_cb(f'  🏁 "Chọn tất cả" → đã chọn hết đơn, dừng', 'dim')
                    else:
                        force_stop = False
                        log_cb(f'  🔄 "Chọn X đầu tiên" → còn đơn, chạy batch tiếp', 'info')
                    select_all_batches += 1
                    # ── Safety limit: tránh loop vô hạn nếu TikTok thay đổi UI ──
                    if select_all_batches > 20:
                        log_cb('  ⚠ Đã chạy 20 batch "Chọn tất cả" — dừng để tránh loop vô hạn', 'warn')
                        force_stop = True
                else:
                    # ✅ Case 1: KHÔNG có nút "Chọn tất cả" → chỉ có đúng ≤20 đơn thật
                    checked = header_checked
                    log_cb(f'  ✓ Chỉ có {checked} đơn trên trang — không có nút "Chọn tất cả" (Case 1)', 'ok')
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
            page.wait_for_timeout(5000)  # Đợi TikTok UI phản ứng sau khi chọn đơn
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
                # Fallback: có thể checkbox chưa thực sự được tick (false positive từ đếm JS)
                # Thử tick tay từng checkbox rồi tìm lại nút
                log_cb('  ⚠ Không tìm thấy nút — thử tick tay từng checkbox...', 'warn')
                cbs = page.query_selector_all('td.col-checkbox label.p-checkbox')
                retry_checked = 0
                for cb in cbs:
                    try:
                        cb.click()
                        retry_checked += 1
                        page.wait_for_timeout(100)
                    except: pass
                page.wait_for_timeout(1500)
                # Tìm lại ship button
                for btn_text in ['Sắp xếp vận chuyển và in', 'Arrange shipment and print', 'Sắp xếp vận chuyển']:
                    try:
                        btn = page.locator(f'button:has-text("{btn_text}")').first
                        if btn.count() > 0 and btn.is_visible(timeout=2000): ship_btn = btn; break
                    except: pass
                if not ship_btn:
                    all_btns = page.locator('button').all()
                    for b in all_btns:
                        try:
                            txt = b.inner_text().strip().lower()
                            if 'vận chuyển' in txt and 'in' in txt: ship_btn = b; break
                        except: pass
                if ship_btn:
                    checked = retry_checked
                    log_cb(f'  ✓ Đã tick tay {checked} đơn và tìm thấy nút', 'ok')
            if not ship_btn:
                log_cb('  ✗ KHÔNG TÌM THẤY nút "Sắp xếp vận chuyển và in"!', 'err')
                try:
                    ss = str(Path(output_dir) / f'debug_no_ship_btn_batch{batch_num}.png')
                    page.screenshot(path=ss); log_cb(f'  📸 Screenshot: {ss}', 'info')
                except: pass
                total_printed += checked; break
            ship_btn.click()
            log_cb('  ✓ Đã bấm "Sắp xếp vận chuyển và in"', 'ok')
            page.wait_for_timeout(5000)

            # ── Popup "cùng địa chỉ" (TikTok mới thêm) ──
            # Nếu có đơn cùng địa chỉ → popup hỏi gộp đơn.
            # Phải chọn "Tiếp tục mà không kết hợp" trước khi tới bước "Tiếp theo".
            state_cb('printing', f'Batch {batch_num}: Kiểm tra popup "cùng địa chỉ"...')
            khong_ket_hop_btn = None
            for _ in range(15):
                for sel_text in ['Tiếp tục mà không kết hợp', 'Continue without combining',
                                 'Không kết hợp', 'Do not combine',
                                 'Tiếp tục', 'Continue']:
                    try:
                        btn = page.locator(f'button:has-text("{sel_text}")').first
                        if btn.count() > 0 and btn.is_visible(timeout=500):
                            khong_ket_hop_btn = btn
                            break
                    except Exception:
                        pass
                if khong_ket_hop_btn:
                    break
                page.wait_for_timeout(1000)

            if khong_ket_hop_btn:
                khong_ket_hop_btn.click(timeout=5000)
                log_cb('  ✓ Đã bấm "Tiếp tục mà không kết hợp" (popup cùng địa chỉ)', 'ok')
                page.wait_for_timeout(3000)
            else:
                log_cb('  ℹ Không có popup cùng địa chỉ — tiếp tục...', 'dim')

            # ── Helper: tìm & click nút "Tiếp tục" (popup xác nhận trung gian) ──
            # Có thể bị overlay (vd "Phí vận chuyển") chặn → thử nhiều cách, không lỗi
            def _try_click_tieptuc(reason=''):
                for _ in range(5):
                    for sel_text in ['Tiếp tục', 'Continue', 'Xác nhận', 'Confirm', 'OK']:
                        try:
                            candidates = page.locator(f'button:has-text("{sel_text}")').all()
                            for btn in candidates:
                                try:
                                    txt = btn.inner_text().strip().lower()
                                    if not btn.is_visible(timeout=300):
                                        continue
                                    if 'không kết hợp' in txt or 'without combining' in txt:
                                        continue
                                    if 'do not combine' in txt:
                                        continue
                                    if 'kết hợp' in txt and 'không' not in txt:
                                        continue  # Bỏ qua nút "chấp nhận tất cả X kết hợp và tiếp tục"
                                    if 'combine' in txt and 'without' not in txt and 'do not' not in txt:
                                        continue  # Bỏ qua nút "accept all X combinations and continue"
                                    if 'tiếp theo' in txt or 'next' in txt:
                                        continue
                                    if sel_text.lower() in txt:
                                        # Thử click bình thường
                                        try:
                                            btn.click(timeout=2000)
                                            log_cb(f'  ✓ Đã bấm "{txt[:40]}" {reason}', 'ok')
                                            page.wait_for_timeout(2000)
                                            return True
                                        except Exception:
                                            # Bị chặn → thử force click
                                            try:
                                                btn.click(force=True, timeout=2000)
                                                log_cb(f'  ✓ Đã force-click "{txt[:40]}" {reason}', 'ok')
                                                page.wait_for_timeout(2000)
                                                return True
                                            except Exception:
                                                # Force cũng fail → thử JS dispatchEvent
                                                try:
                                                    btn.dispatch_event('click')
                                                    log_cb(f'  ✓ Đã JS-click "{txt[:40]}" {reason}', 'ok')
                                                    page.wait_for_timeout(2000)
                                                    return True
                                                except Exception:
                                                    pass
                                except Exception:
                                    pass
                        except Exception:
                            pass
                    page.wait_for_timeout(800)
                return False

            # ── Popup xác nhận trung gian (TikTok mới thêm sau bước "cùng địa chỉ") ──
            state_cb('printing', f'Batch {batch_num}: Kiểm tra popup xác nhận trung gian...')
            clicked_tieptuc = _try_click_tieptuc('(popup xác nhận trung gian)')
            if not clicked_tieptuc:
                log_cb('  ℹ Không có popup xác nhận trung gian — tiếp tục...', 'dim')

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
                page.wait_for_timeout(4000)
                # Thử click "Tiếp tục" lần nữa (phòng overlay "Phí vận chuyển" đã biến mất)
                _try_click_tieptuc('(sau "Tiếp theo")')
            else:
                log_cb('  ⏭ Không thấy nút "Tiếp theo" — vẫn kiểm tra hộp thoại chọn loại chứng từ/nhãn', 'dim')
                # 📸 Chụp UI thật để đối chiếu nếu bug vẫn còn
                try:
                    ss = str(Path(output_dir) / f'debug_no_tieptheo_batch{batch_num}.png')
                    page.screenshot(path=ss)
                    log_cb(f'  📸 Screenshot: {ss}', 'dim')
                except: pass

            # ── Chọn loại chứng từ / nhãn vận chuyển (LUÔN chạy, KHÔNG gate theo "Tiếp theo") ──
            # TikTok CÓ LÚC hiện hộp thoại chọn này, CÓ LÚC không → phải poll + delay cho web
            # kịp mở dialog, có checkbox thì tick, không có thì bỏ qua rồi vẫn đi tải xuống.
            state_cb('printing', f'Batch {batch_num}: Chọn loại chứng từ/nhãn (nếu có)...')
            page.wait_for_timeout(3000)  # chờ dialog tải xong sau khi bấm
            doc_labels = ['Danh sách đóng gói', 'Danh sách lấy hàng',
                          'Shipping label', 'Nhãn vận chuyển']
            found_any_doc = False
            for _ in range(15):  # poll tối đa 15s cho các checkbox xuất hiện (dialog load chậm)
                for doc_label in doc_labels:
                    try:
                        lbl = page.locator('label').filter(has_text=doc_label).first
                        if lbl.count() > 0:
                            inp = lbl.locator('input')
                            if inp.count() > 0:
                                found_any_doc = True
                                if not inp.is_checked():
                                    lbl.click(); page.wait_for_timeout(500)
                                    log_cb(f'  ✓ Đã tick: {doc_label}', 'ok')
                    except: pass
                if found_any_doc:
                    break  # đã thấy dialog → đủ điều kiện, không cần poll tiếp
                page.wait_for_timeout(1000)
            if not found_any_doc:
                log_cb('  ℹ Không có hộp thoại chọn loại chứng từ/nhãn — bỏ qua', 'dim')
            page.wait_for_timeout(2000)
            # Thử click "Tiếp tục" lần nữa trước khi in
            _try_click_tieptuc('(trước "In nhãn ngay")')

            # ── "In nhãn ngay" là TÙY CHỌN: có thì bấm, không có thì vẫn đi tải ──
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
            else: log_cb('  ⚠ Không tìm thấy nút "In nhãn ngay" — đi tải với mặc định hệ thống', 'warn')

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

            # ── Click "chịu khó" — normal click hay bị treo nếu có overlay vô hình
            # chặn pointer event (vd toast/backdrop đang fade out) → fallback force
            # click rồi JS dispatch, tránh làm crash toàn bộ automation.
            def _robust_click(btn, desc=''):
                try:
                    btn.click(timeout=5000)
                    return True
                except Exception as e1:
                    log_cb(f'  ⚠ Click thường thất bại {desc} ({e1}) — thử force click...', 'warn')
                try:
                    btn.click(timeout=3000, force=True)
                    return True
                except Exception as e2:
                    log_cb(f'  ⚠ Force click thất bại {desc} ({e2}) — thử JS click...', 'warn')
                try:
                    btn.dispatch_event('click')
                    return True
                except Exception as e3:
                    log_cb(f'  ✗ Không thể click {desc}: {e3}', 'err')
                    # ── Chẩn đoán: cả 3 cách đều fail — chụp lại phần tử đang thực
                    # sự nằm ở đúng toạ độ nút để xác nhận có bị che hay không ──
                    try:
                        box = btn.bounding_box()
                        if box:
                            cx = box['x'] + box['width'] / 2
                            cy = box['y'] + box['height'] / 2
                            blocker = page.evaluate(
                                "([x, y]) => { const el = document.elementFromPoint(x, y); "
                                "return el ? (el.tagName + '.' + (el.className || '') + ' | text=' + "
                                "(el.innerText || '').slice(0, 40)) : 'none'; }",
                                [cx, cy])
                            log_cb(f'  🔍 Phần tử thực tế tại vị trí nút: {blocker}', 'dim')
                        ss = str(Path(output_dir) / f'debug_click_fail_batch{batch_num}_{datetime.now().strftime("%H-%M-%S")}.png')
                        page.screenshot(path=ss)
                        log_cb(f'  📸 Screenshot chẩn đoán: {ss}', 'dim')
                    except Exception as diag_err:
                        log_cb(f'  ⚠ Không chụp được chẩn đoán: {diag_err}', 'dim')
                    return False

            # ── Download với retry (tối đa 3 lần) ──
            MAX_DOWNLOAD_RETRIES = 3
            download_ok = False
            for retry_attempt in range(MAX_DOWNLOAD_RETRIES):
                if retry_attempt > 0:
                    log_cb(f'  🔄 Retry download lần {retry_attempt+1}/{MAX_DOWNLOAD_RETRIES}...', 'warn')
                    page.wait_for_timeout(3000)
                    # Thử click lại nút tải xuống nếu popup còn hiển thị
                    retry_btn = None
                    for sel in ['button:has-text("Tải xuống tất cả")', 'button:has-text("Download all")',
                                'button:has-text("Tải xuống")']:
                        try:
                            btn = page.locator(sel).first
                            if btn.count() > 0 and btn.is_visible(timeout=2000):
                                retry_btn = btn; break
                        except: pass
                    if retry_btn:
                        if _robust_click(retry_btn, '(retry "Tải xuống tất cả")'):
                            log_cb('  ✓ Đã click lại "Tải xuống tất cả"', 'ok')
                        else:
                            continue
                    else:
                        log_cb('  ⚠ Popup tải xuống đã biến mất — không thể retry', 'warn')
                        break

                downloaded_files = []
                def on_download(dl):
                    carrier_prefix = carrier.replace(' ', '_').replace('&', 'n') + '_' if carrier else ''
                    base_name = dl.suggested_filename or f'PDF_goc_TTS_batch{batch_num}_{len(downloaded_files)}_{datetime.now().strftime("%m-%d_%H-%M-%S")}.pdf'
                    suggested = carrier_prefix + base_name if carrier else base_name
                    bp = str(Path(output_dir) / suggested)
                    try:
                        dl.save_as(bp)
                        downloaded_files.append(bp)
                        log_cb(f'  💾 Đã tải: {Path(bp).name}', 'ok')
                    except Exception as save_err:
                        log_cb(f'  ⚠ Lỗi lưu file: {save_err}', 'warn')

                page.on('download', on_download)
                if retry_attempt == 0:
                    if _robust_click(taixuong_btn, '("Tải xuống tất cả")'):
                        log_cb('  ✓ Đã bấm "Tải xuống tất cả"', 'ok')
                    else:
                        page.remove_listener('download', on_download)
                        continue
                # Chờ download hoàn tất: tối đa 3 phút mỗi lần retry
                idle_ticks = 0
                for _ in range(18):  # 18 × 10s = 3 phút tối đa
                    page.wait_for_timeout(10000)
                    if downloaded_files:
                        prev_count = len(downloaded_files)
                        page.wait_for_timeout(3000)
                        if len(downloaded_files) == prev_count:
                            idle_ticks += 1
                            if idle_ticks >= 2:  # 26 giây không có download mới → xong
                                log_cb(f'  ✅ Download hoàn tất ({len(downloaded_files)} file)', 'ok')
                                break
                        else:
                            idle_ticks = 0  # reset, vẫn còn download mới
                page.remove_listener('download', on_download)

                # ── VALIDATE file đã tải ──
                valid_files = []
                for f in downloaded_files:
                    try:
                        if Path(f).exists() and Path(f).stat().st_size > 0:
                            valid_files.append(f)
                        else:
                            log_cb(f'  ⚠ File lỗi (0 bytes hoặc thiếu): {Path(f).name}', 'warn')
                    except Exception as ve:
                        log_cb(f'  ⚠ Không kiểm tra được file: {Path(f).name} - {ve}', 'warn')

                if valid_files:
                    pdf_files.extend(valid_files)
                    log_cb(f'  📥 Đã tải {len(valid_files)} file hợp lệ', 'info')
                    download_ok = True
                    break
                else:
                    log_cb(f'  ⚠ Không có file hợp lệ trong lần tải này', 'warn')

            if not download_ok:
                log_cb('  ✗ Download thất bại sau các lần retry — bỏ qua batch này', 'err')
                total_printed += checked; break

            total_printed += checked
            log_cb(f'  📊 Tiến độ: {total_printed}/{target} đơn, {len(pdf_files)} file PDF', 'info')

            if force_stop:
                log_cb('  ✓ Đã in hết đơn hiện có — hoàn thành.', 'ok'); break
            if checked < 20:
                log_cb(f'  ✓ Batch này chỉ có {checked} đơn (< 20) — hoàn thành.', 'ok'); break
            if total_printed < target and total_avail > 0:
                page.wait_for_timeout(2000)

        return pdf_files, playwright, browser
    except Exception as e:
        log_cb(f'  ✗ Lỗi: {e}', 'err')
        return pdf_files, playwright, browser

# ============================================================
# SEND ORDER IDS TO API
# ============================================================
def _send_order_ids_to_api(order_ids, log_cb, url=None):
    """Gửi danh sách Order ID đến API qua HTTP POST (text/plain)."""
    if not order_ids:
        return False
    if url is None:
        url = API_IDS_URL
    try:
        import urllib.request
        body = ','.join(order_ids) + ','
        # Đảm bảo không có dấu phẩy kép nếu danh sách rỗng
        body = body if body != ',' else ''
        data = body.encode('utf-8')
        req = urllib.request.Request(url, data=data, method='POST')
        req.add_header('Content-Type', 'text/plain')
        with urllib.request.urlopen(req, timeout=30) as resp:
            status = resp.status
            if 200 <= status < 300:
                log_cb(f'  📤 Đã gửi {len(order_ids)} Order ID → API ({status})', 'ok')
                return True
            else:
                log_cb(f'  ⚠ API trả về {status}', 'warn')
                return False
    except Exception as e:
        log_cb(f'  ⚠ Lỗi gửi Order ID lên API: {e}', 'warn')
        return False


# ============================================================
# CARRIER DETECTION (tên file)
# ============================================================
CARRIER_DETECT_RULES = [
    ('ghn', 'GHN'),
    ('jnt_cargo', 'J&T Cargo VN'),  # phải trước 'jnt' — 'JnT_Cargo_VN_...' chứa 'jnt'
    ('j&t cargo', 'J&T Cargo VN'),
    ('vietnam post', 'VietNam Post'),
    ('best express', 'Best Express'),
    ('jnt', 'J&T'),
    ('j&t', 'J&T'),
    ('vietnam', 'VietNam Post'),
    ('vnp', 'VietNam Post'),
    ('best', 'Best Express'),
    ('viettel', 'Viettel Post'),
    ('jtc', 'J&T Cargo VN'),
]


def _detect_carrier_from_filename(fname: str) -> str:
    """Nhận diện hãng vận chuyển từ tên file (GHN_..., JnT_..., ...). Trả về '' nếu không rõ."""
    fn = fname.lower()
    for key, label in CARRIER_DETECT_RULES:
        if key in fn:
            return label
    return ''


# ============================================================
# CALCULATOR
# ============================================================
def run_calculator(pdf_paths, output_dir, master_path, retail_path, template_path, log_cb, carrier='', send_order_ids=True):
    try:
        from calculator import process_all  # Lazy import — chỉ load khi chạy calculator
    except ImportError:
        log_cb('✗ Calculator không khả dụng (thiếu module calculator)', 'err'); return []
    if not Path(master_path).exists(): log_cb(f'✗ Không tìm thấy master_data: {master_path}', 'err'); return []
    if not Path(template_path).exists(): log_cb(f'✗ Không tìm thấy template: {template_path}', 'err'); return []
    out_dir = str(output_dir)
    for p in pdf_paths:
        shutil.copy2(p, str(UPLOAD_DIR / Path(p).name))
    try:
        results = process_all(pdf_paths, out_dir, master_path, retail_path, carrier, template_path=template_path)
        for r in results:
            log_cb(f'  ✓ {r["rows"]} dòng | Qty={r["tong_qty"]} | Sold={r["tong_sold"]} | Promo={r["tong_promo"]}', 'ok')
            for key, fb in r['files'].items():
                src, dst = Path(fb), Path(out_dir) / Path(fb).name
                if src != dst and src.exists(): shutil.copy2(str(src), str(dst)); r['files'][key] = str(dst)
            # Gửi Order ID lên API nếu được bật
            if send_order_ids:
                order_id_txt = r['files'].get('order_id_txt')
                if order_id_txt and Path(order_id_txt).exists():
                    try:
                        with open(order_id_txt, 'r', encoding='utf-8') as f:
                            lines = [l.strip() for l in f if l.strip() and not l.startswith('#')]
                        if lines:
                            _send_order_ids_to_api(lines, log_cb)
                    except Exception as e:
                        log_cb(f'  ⚠ Lỗi đọc Order ID để gửi API: {e}', 'warn')
            else:
                log_cb('  ℹ Bỏ qua gửi Order ID (đã tắt trong tab Test)', 'dim')
        return results
    except Exception as e:
        import traceback
        log_cb(f'  ✗ Lỗi: {e}', 'err')
        log_cb(f'  📋 Traceback: {traceback.format_exc()}', 'dim')
        return []

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
        # ── Dọn dẹp browser cũ từ job trước (nếu có) ──
        if self._browser is not None:
            try:
                self._browser.close()
            except Exception:
                pass
            self._browser = None
        if self._playwright is not None:
            import time as _time_cleanup
            _time_cleanup.sleep(0.3)
            try:
                self._playwright.stop()
            except Exception:
                pass
            self._playwright = None

        self._running = True
        self._stop_event.clear()

        all_pdf_paths = []
        all_results = []

        pw = None
        br = None

        try:
            cookie = config['cookie']
            out_dir = config['output_dir']  # Đã có sẵn date subfolder từ main thread
            master = config['master']
            retail = config['retail']
            template = config['template']
            auto_print = config['auto_print']
            printer = config['printer']
            test_mode = config['test_mode']
            exclude_pre_orders = config.get('exclude_pre_orders', True)
            batch_size = config.get('batch_size', 0)
            pdf_settings = config.get('pdf_settings', 'paper=A4')
            carriers = config['carriers']

            os.makedirs(out_dir, exist_ok=True)

            for carrier, count in carriers:
                # ── Skip completed carriers khi resume ──
                if self._stop_event.is_set():
                    self.log_message.emit('warn', 'Đã dừng theo yêu cầu.');
                    break

                carrier_display = carrier if carrier else 'tất cả'
                count_display = f'{count}' if count > 0 else 'tất cả'
                self.log_message.emit('info', f'📥 Tải PDF [{carrier_display}] ({count_display} đơn)...')

                pdf_paths, pw2, br2 = run_automation(
                    cookie, out_dir, count,
                    lambda m, t='': self.log_message.emit(t, m),
                    lambda s, m: self.state_changed.emit(s, m),
                    self._stop_event,
                    existing_playwright=pw, existing_browser=br,
                    carrier=carrier, test_mode=test_mode,
                    exclude_pre_orders=exclude_pre_orders)

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

                # In shipping label TRƯỚC, rồi mới tính toán
                if auto_print and pdf_paths:
                    self.log_message.emit('info', f'🖨️ [{carrier_display}]: In shipping label...')
                    for p in pdf_paths:
                        if self._stop_event.is_set():
                            self.log_message.emit('warn', '⏹ Đã dừng in — hủy các file còn lại.');
                            break
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
                if pdf_paths and not self._stop_event.is_set():
                    self.log_message.emit('info', f'📊 Đang tính bill [{carrier_display}]...')
                    carrier_results = run_calculator(
                        pdf_paths, out_dir, master, retail, template,
                        lambda m, t='': self.log_message.emit(t, m),
                        carrier=carrier)
                    for r in carrier_results:
                        for key, lbl in [('xlsx_report', '📊')]:
                            fp = r['files'].get(key)
                            if fp and Path(fp).exists():
                                self.log_message.emit('info', f'{lbl} {Path(fp).name}')
                                self.result_file.emit(f'{lbl} {Path(fp).name}')
                    all_results.extend(carrier_results)

                # In báo cáo sau khi tính toán
                if auto_print and carrier_results and not self._stop_event.is_set():
                    for r in carrier_results:
                        if self._stop_event.is_set():
                            self.log_message.emit('warn', '⏹ Đã dừng in báo cáo.');
                            break
                        fp = r['files'].get('pdf_report') or r['files'].get('xlsx_report')
                        if fp and Path(fp).exists():
                            try:
                                self.log_message.emit('info', f'  🖨️ In báo cáo: {Path(fp).name}')
                                _print_file(fp, printer, pdf_settings=pdf_settings, batch_size=batch_size,
                                            log_cb=lambda m, t='': self.log_message.emit(t, m))
                                self.log_message.emit('ok', f'  ✓ Đã in báo cáo: {Path(fp).name}')
                            except Exception as e:
                                self.log_message.emit('err', f'  ✗ Lỗi in báo cáo: {e}')

            # ── Tất cả carriers hoàn thành ──
            self.log_message.emit('bold_ok', '🏁 HOÀN THÀNH!')
            self.state_changed.emit('done', f'✅ Hoàn thành lúc {datetime.now().strftime("%H:%M:%S")}')
            try:
                if br: br.close()
            except Exception: pass
            try:
                if pw: pw.stop()
            except Exception: pass
            self._playwright = None
            self._browser = None
            self.job_completed.emit({
                'playwright': None, 'browser': None,
                'pdf_paths': all_pdf_paths, 'results': all_results,
                'output_dir': out_dir,
            })
        except Exception as e:
            self.log_message.emit('err', f'✗ Lỗi: {e}')
            self.state_changed.emit('error', f'✗ {e}')
            try:
                if self._browser: self._browser.close()
            except Exception: pass
            try:
                if self._playwright: self._playwright.stop()
            except Exception: pass
            self._playwright = None
            self._browser = None
            self.job_completed.emit({})
        finally:
            self._running = False

    @Slot()
    def stop_job(self): self._stop_event.set()

    def shutdown_browser(self):
        """Force-close browser NGAY LẬP TỨC (thread-safe, gọi từ main thread)."""
        self._stop_event.set()
        br = self._browser
        if br is not None:
            try:
                br.close()
            except Exception:
                pass
        # Không stop playwright ở đây — chỉ close browser.
        # stop playwright gọi vào native code dễ crash nếu đang dở việc.

    @Slot()
    def shutdown(self):
        self._stop_event.set()
        try:
            if self._browser: self._browser.close()
        except: pass
        try:
            if self._playwright: self._playwright.stop()
        except: pass

_foxit_exe_cache = None


def _find_foxit_exe():
    """Tìm Foxit PDF Reader — dùng XPS Print Path, spool nhẹ ~15MB."""
    global _foxit_exe_cache
    if _foxit_exe_cache is not None:
        return _foxit_exe_cache or ''
    foxit_paths = [
        r'C:\Program Files (x86)\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe',
        r'C:\Program Files\Foxit Software\Foxit PDF Reader\FoxitPDFReader.exe',
        r'C:\Program Files (x86)\Foxit Software\Foxit PhantomPDF\FoxitPhantomPDF.exe',
        r'C:\Program Files\Foxit Software\Foxit PhantomPDF\FoxitPhantomPDF.exe',
    ]
    # Tìm thêm trong thư mục con của Foxit Software (1 cấp)
    for base in [r'C:\Program Files (x86)\Foxit Software', r'C:\Program Files\Foxit Software']:
        try:
            if os.path.isdir(base):
                # Quét file trực tiếp trong base
                for f in os.listdir(base):
                    fp_full = os.path.join(base, f)
                    if os.path.isfile(fp_full) and f.lower().startswith('foxit') and f.lower().endswith('.exe'):
                        foxit_paths.append(fp_full)
                # Quét 1 cấp thư mục con
                for sub in os.listdir(base):
                    sub_path = os.path.join(base, sub)
                    if os.path.isdir(sub_path):
                        try:
                            for f in os.listdir(sub_path):
                                if f.lower().startswith('foxit') and f.lower().endswith('.exe'):
                                    foxit_paths.append(os.path.join(sub_path, f))
                        except Exception:
                            pass
        except Exception:
            pass
    for fp in foxit_paths:
        if Path(fp).exists():
            _foxit_exe_cache = fp
            return fp
    _foxit_exe_cache = ''
    return ''


def _check_print_errors(printer_name, doc_name_hint='', log_cb=None, timeout=60):
    """Poll print queue để phát hiện lỗi máy in (kẹt giấy, hết mực...).
    Đợi đến khi job Complete hoặc Error thì trả về.
    timeout: số giây tối đa chờ (mặc định 60s).
    Lưu ý: nếu job chưa từng xuất hiện trong queue sau 15s → coi như đã in xong quá nhanh."""
    import subprocess as _sp, os as _os
    deadline = __import__('time').time() + timeout
    last_status = ''
    error_reported = False
    never_seen_deadline = __import__('time').time() + 15  # 15s đầu phải thấy job, nếu không → exit sớm
    while __import__('time').time() < deadline:
        __import__('time').sleep(5)
        result = _sp.run(['powershell', '-NoProfile', '-WindowStyle', 'Hidden', '-Command',
            f"$jobs = Get-PrintJob -PrinterName '{printer_name}' -ErrorAction SilentlyContinue"
            + (f" | Where-Object {{ $_.DocumentName -like '*{doc_name_hint[:30].replace(chr(39), '').replace(chr(34), '')}*' }}" if doc_name_hint else "")
            + " | Select-Object JobStatus | ConvertTo-Json -Compress"
        ], capture_output=True, text=True,
           creationflags=_sp.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
        status_raw = result.stdout.strip()
        status = ''
        if status_raw:
            try:
                import json as _json
                job_info = _json.loads(status_raw)
                if isinstance(job_info, list):
                    job_info = job_info[0] if job_info else {}
                status = job_info.get('JobStatus', '')
            except Exception:
                status = status_raw

        if status and status != last_status:
            if log_cb: log_cb(f'  🖨️ Trạng thái: {status}', 'dim')
            last_status = status

        if status:
            if 'Complete' in status and 'Printing' not in status and 'Spooling' not in status:
                if log_cb: log_cb(f'  ✅ Job in đã hoàn thành', 'dim')
                return True

            error_msgs = {
                'PaperJam': '🛑 KẸT GIẤY! Hãy gỡ giấy kẹt rồi nhấn nút trên máy in.',
                'PaperOut': '📄 HẾT GIẤY! Hãy nạp thêm giấy vào khay.',
                'TonerLow': '⚠ SẮP HẾT MỰC! Chuẩn bị thay mực.',
                'NoToner': '🖌 HẾT MỰC! Cần thay cartridge mực.',
                'Offline': '🔌 MÁY IN MẤT KẾT NỐI! Kiểm tra cáp/WiFi.',
                'Paused': '⏸ MÁY IN ĐANG TẠM DỪNG! Kiểm tra nút trên máy.',
                'DoorOpen': '🚪 NẮP MÁY IN ĐANG MỞ! Đóng nắp lại.',
                'OutputFull': '📦 KHAY RA ĐẦY! Lấy giấy đã in ra.',
            }
            for code, msg in error_msgs.items():
                if code in status and not error_reported:
                    if log_cb: log_cb(f'  {msg}', 'err')
                    error_reported = True
                    break

            if 'Error' in status and not error_reported:
                if log_cb: log_cb(f'  ⚠ Lỗi máy in: {status}', 'err')
                return False
        else:
            # Job đã biến mất khỏi queue (đã in xong và được xóa)
            if last_status:
                if log_cb: log_cb(f'  ✅ Job in đã rời queue (đã in xong)', 'dim')
                return True
            # Neu da qua 15s ma chua bao gio thay job → may in xu ly qua nhanh, coi nhu xong
            if __import__('time').time() > never_seen_deadline:
                if log_cb: log_cb(f'  ⚡ Job in qua nhanh, khong thay trong queue — coi nhu da in xong', 'dim')
                return True

    # Hết timeout — có thể job vẫn đang in, không chặn tiến trình
    if log_cb: log_cb(f'  ⚠ Hết {timeout}s chờ — tiếp tục (job có thể vẫn đang in)', 'warn')
    return True


def _wait_print_queue(printer_name, max_jobs=2, timeout=30):
    """Đợi hàng đợi máy in ≤ max_jobs rồi mới gửi job mới (tránh quá tải RAM máy in).
    timeout: số giây tối đa chờ (mặc định 30s)."""
    import subprocess as _sp, time as _t
    deadline = _t.time() + timeout
    waited = False
    while _t.time() < deadline:
        result = _sp.run(['powershell', '-NoProfile', '-WindowStyle', 'Hidden', '-Command',
            f"@(Get-PrintJob -PrinterName '{printer_name}' -ErrorAction SilentlyContinue).Count"
        ], capture_output=True, text=True, creationflags=_sp.CREATE_NO_WINDOW if sys.platform == 'win32' else 0)
        try:
            count = int(result.stdout.strip())
        except ValueError:
            count = 0
        if count <= max_jobs:
            if waited:
                print(f'   ✓ Hàng đợi đã trống (sau ~{int(_t.time() - (deadline - timeout))}s)')
            return
        if not waited:
            waited = True
            print(f'   ⏳ Hàng đợi có {count} job(s) — đợi giảm xuống ≤{max_jobs}...')
        _t.sleep(1)  # kiểm tra mỗi 1 giây
    # Nếu hết timeout mà vẫn còn job → log cảnh báo và in tiếp (tránh treo vĩnh viễn)
    print(f'   ⚠ Hết {timeout}s chờ — hàng đợi vẫn còn job, in tiếp...')


def _print_file(file_path, printer_name, pdf_settings='paper=A4', log_cb=None, batch_size=0):
    import subprocess, os as _os
    fp = str(file_path)

    try:
        if fp.lower().endswith('.pdf'):
            foxit_exe = _find_foxit_exe()
            if not foxit_exe:
                raise RuntimeError(
                    'Không tìm thấy Foxit PDF Reader. '
                    'Vui lòng cài Foxit PDF Reader để in file PDF.'
                )

            print_path = fp
            temp_merged = None
            # Chỉ merge 2-up file shipping label, không merge file báo cáo
            fname_lower = Path(fp).name.lower()
            if 'shipping' in fname_lower or 'vận chuyển' in fname_lower:
                try:
                    if log_cb: log_cb(f'  📐 Đang merge 2-up: {Path(fp).name}...', 'dim')
                    temp_merged = _merge_pdf_2up(fp)
                    if temp_merged: print_path = temp_merged
                    if log_cb: log_cb(f'  ✓ Merge 2-up hoàn tất', 'dim')
                except Exception as e:
                    if log_cb: log_cb(f'  ⚠ Merge 2-up lỗi ({e}) — in file gốc', 'warn')

            if foxit_exe:
                # ── Đọc số trang để quyết định batch splitting ──
                try:
                    from pypdf import PdfReader as _PdfReader, PdfWriter as _PdfWriter
                    reader = _PdfReader(print_path)
                    total_pages = len(reader.pages)
                except Exception:
                    reader = None
                    total_pages = 0

                # ── Batch splitting ──
                if batch_size > 0 and reader and total_pages > batch_size:
                    total_batches = (total_pages + batch_size - 1) // batch_size
                    if log_cb: log_cb(f'  📦 Foxit: Chia {total_pages} tờ → {total_batches} batch ({batch_size} tờ/batch)', 'info')
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

                        if log_cb: log_cb(f'  ⏳ Đợi hàng đợi máy in trống...', 'dim')
                        _wait_print_queue(printer_name, max_jobs=0)
                        if log_cb: log_cb(f'  ▶ Gửi lệnh in qua Foxit (batch {batch_num})...', 'info')
                        cmd = [foxit_exe, '/t', batch_path, printer_name]
                        result = subprocess.run(cmd, check=False, timeout=3600)
                        if log_cb: log_cb(f'  ✓ Foxit batch {batch_num} đã thoát (exit code: {result.returncode})', 'dim')
                        if result.returncode != 0:
                            raise RuntimeError(f'Foxit batch {batch_num} exit code: {result.returncode}')
                        if log_cb: log_cb(f'  🔍 Đang kiểm tra trạng thái in batch {batch_num}...', 'dim')
                        _check_print_errors(printer_name, doc_name_hint=os.path.basename(batch_path), log_cb=log_cb)

                        if log_cb: log_cb(f'  ✅ Batch {batch_num}/{total_batches} đã in xong', 'ok')
                        try: _os.remove(batch_path)
                        except: pass

                else:
                    # ── In thẳng không batch ──
                    if log_cb: log_cb(f'  ⏳ Đợi hàng đợi máy in trống (max_jobs=0)...', 'dim')
                    _wait_print_queue(printer_name, max_jobs=0)
                    if log_cb: log_cb(f'  ▶ Gửi lệnh in qua Foxit: {os.path.basename(print_path)}', 'info')
                    cmd = [foxit_exe, '/t', print_path, printer_name]
                    result = subprocess.run(cmd, check=False, timeout=3600)
                    if log_cb: log_cb(f'  ✓ Foxit đã thoát (exit code: {result.returncode})', 'dim')
                    if result.returncode != 0:
                        raise RuntimeError(f'Foxit exit code: {result.returncode}')
                    if log_cb: log_cb(f'  🔍 Đang kiểm tra trạng thái in...', 'dim')
                    _check_print_errors(printer_name, doc_name_hint=os.path.basename(print_path), log_cb=log_cb)
                    if log_cb: log_cb(f'  ✅ In hoàn tất', 'ok')
            if temp_merged:
                try:
                    _os.remove(temp_merged)
                    if log_cb: log_cb(f'  🗑 Đã xóa file tạm merge 2-up', 'dim')
                except Exception:
                    pass
            return

        if fp.lower().endswith('.xlsx') or fp.lower().endswith('.xls'):
            import pythoncom, win32com.client, time as _t_excel
            pythoncom.CoInitialize()
            excel = None
            try:
                # Đợi queue trống + delay cứng để đảm bảo 2 job không bị gộp
                _wait_print_queue(printer_name)
                _t_excel.sleep(1)
                excel = win32com.client.Dispatch("Excel.Application")
                excel.Visible = False
                workbook = excel.Workbooks.Open(_os.path.abspath(fp))
                workbook.PrintOut(ActivePrinter=printer_name, FitToPagesWide=1, FitToPagesTall=False)
                # Đợi Excel spool xong job ra queue rồi mới đóng
                _t_excel.sleep(2)
                workbook.Close(False)
            finally:
                try:
                    if excel: excel.Quit()
                except: pass
                try:
                    pythoncom.CoUninitialize()
                except: pass
            # Đợi job đã chắc chắn vào queue
            _t_excel.sleep(1)
            return
    except Exception:
        raise

def _merge_pdf_2up(pdf_path):
    import os as _os
    try: from pypdf import PdfReader, PdfWriter, PageObject, Transformation
    except ImportError: return None
    try:
        reader = PdfReader(pdf_path)
        if len(reader.pages) < 1: return None
        canvas_w, canvas_h = 842, 595
        margin_left, margin_top, gap, scale_factor = 0, 28, -14, 0.70
        writer = PdfWriter()
        avail_w = canvas_w - 2*margin_left - gap
        half_w = avail_w / 2
        # Nếu lẻ 1 trang → vẫn merge 1 mình lên A4 ngang để scale nhỏ, tránh bị to nguyên tờ
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
        local_template = BASE_DIR / "Bảng thống kê hàng.xlsx"
        if local_template.exists(): self._template_real = str(local_template)
        elif TEMPLATE_DEFAULT.exists(): self._template_real = str(TEMPLATE_DEFAULT)
        else: self._template_real = ""

        self.running = False
        self.scheduler_active = False
        self.result_files = []
        self._stop_event = threading.Event()

        self._sched_mode = "weekly"
        self._sched_interval_hours = 1
        self._sched_weekly_config: dict[str, dict[int, list[tuple[int, int]]]] = {}  # {"": {day_idx: [(h,m)]}, "ghn": {...}, ...}
        self._sched_carrier_custom: dict[str, bool] = {}  # {"ghn": True, "jt": False, ...} — carrier nào dùng lịch riêng
        self._sched_next_run = None
        self._sched_last_run = None

        # Weekly scheduler widget refs (assigned in _build_schedule_tab)
        self.weekly_rb = None
        self.weekly_panel = None
        self.weekly_day_checkboxes: dict[int, QCheckBox] = {}  # Tab "Tất cả"
        self.weekly_day_time_edits: dict[int, QLineEdit] = {}  # Tab "Tất cả"
        self.weekly_master_time_edit = None
        # Per-carrier schedule widgets
        self._carrier_day_checkboxes: dict[str, dict[int, QCheckBox]] = {}  # {"ghn": {0: cb, 1: cb, ...}, ...}
        self._carrier_day_time_edits: dict[str, dict[int, QLineEdit]] = {}  # {"ghn": {0: te, 1: te, ...}, ...}
        self._carrier_use_custom_cb: dict[str, QCheckBox] = {}  # {"ghn": cb, ...} — checkbox "Dùng lịch riêng"

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

        # ── Nạp cấu hình đã lưu từ lần trước ──
        self._config_path = BASE_DIR / '.tts_config.json'
        self._load_config()

        self._update_cookie_status()
        self._update_master_status()
        self._update_retail_status()
        self._update_template_status()
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
        self.content_stack.addWidget(self._build_aggregate_tab())
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
            "📊 Tổng hợp",
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

        self.template_row = FileRowWidget("📋 Mẫu xuất hàng", "Excel Files (*.xlsx)")
        self.template_row.set_path(self._template_real)
        self.template_row.path_changed.connect(self._on_template_changed)
        gb1_layout.addWidget(self.template_row)

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
            ("GHN:", "ghn"), ("J&T Express:", "jt"),
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
        self.batch_size_spin.setRange(0, 100)
        self.batch_size_spin.setValue(0)
        self.batch_size_spin.setFixedWidth(60)
        self.batch_size_spin.setToolTip("0 = không chia batch. Nhập số >0 để chia nhỏ file in")
        self.batch_size_spin.setSpecialValueText("0 (không chia)")
        paper_row.addWidget(self.batch_size_spin)

        paper_row.addStretch()
        gb_layout.addLayout(paper_row)

        # ── Engine in PDF: Foxit PDF Reader (XPS Print Path, spool ~15MB) ──
        foxit_detected = _find_foxit_exe()

        # Label hiển thị engine in PDF
        engine_label = QLabel("Engine in PDF:")
        engine_label.setFixedWidth(130)
        engine_label.setStyleSheet("font-weight: 600; color: #1E293B; font-size: 10pt;")

        foxit_status = QLabel()
        foxit_status.setWordWrap(True)
        if foxit_detected:
            foxit_status.setText("✅ Foxit PDF Reader — XPS Print Path (spool ~15MB)")
            foxit_status.setStyleSheet("color: #059669; font-size: 9pt; padding: 2px 0;")
        else:
            foxit_status.setText("⚠ Chưa cài Foxit PDF Reader — Vui lòng cài để in file PDF!")
            foxit_status.setStyleSheet("color: #DC2626; font-weight: 600; font-size: 10pt; padding: 2px 0;")

        gb_layout.addWidget(engine_label)
        gb_layout.addWidget(foxit_status)

        # ── Chrome info ──
        chrome_info = QLabel("🌐 Sử dụng Google Chrome có sẵn trên máy")
        chrome_info.setStyleSheet("color: #059669; font-weight: 500; font-size: 13px; margin-top: 8px;")
        gb_layout.addWidget(chrome_info)

        # ── Loại trừ đơn bán trước ──
        self.exclude_pre_orders_cb = QCheckBox("🚫 Loại trừ đơn bán trước (Pre-order) khi tải đơn")
        self.exclude_pre_orders_cb.setChecked(False)
        self.exclude_pre_orders_cb.setToolTip("Bỏ tick nếu bạn MUỐN in cả đơn bán trước.\nTick để bỏ qua đơn bán trước, chỉ in đơn thường.")
        self.exclude_pre_orders_cb.setStyleSheet("color: #64748B; font-weight: 500; font-size: 13px; margin-top: 8px;")
        gb_layout.addWidget(self.exclude_pre_orders_cb)

        # ── Test mode ──
        self.test_mode_cb = QCheckBox("🧪 Bật chế độ Test Mode (Chỉ tải danh sách đơn, KHÔNG thao tác in)")
        self.test_mode_cb.setChecked(False)
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

        self.weekly_rb = QRadioButton("📅 Chạy theo lịch hàng tuần")
        self.weekly_rb.setChecked(True)
        self.weekly_rb.setProperty("mode", "weekly")
        self.sched_button_group.addButton(self.weekly_rb)
        gb_layout.addWidget(self.weekly_rb)

        # Weekly sub-panel — QTabWidget: "Tất cả" + mỗi carrier 1 tab
        self.weekly_panel = QWidget()
        wp_outer = QVBoxLayout(self.weekly_panel)
        wp_outer.setContentsMargins(32, 0, 0, 0)
        wp_outer.setSpacing(8)

        # Quick-fill row
        quick_row = QHBoxLayout()
        quick_row.setSpacing(8)
        quick_row.addWidget(QLabel("Nhập giờ mẫu:"))
        self.weekly_master_time_edit = QLineEdit("07:30, 13:20, 15:10, 16:10, 17:10, 18:10")
        self.weekly_master_time_edit.setFixedWidth(200)
        self.weekly_master_time_edit.setToolTip("Định dạng HH:MM, phân cách bằng dấu phẩy")
        quick_row.addWidget(self.weekly_master_time_edit)
        quick_row.addWidget(QLabel("(phân cách bằng dấu phẩy)"))
        apply_btn = QPushButton("Áp dụng cho tất cả")
        apply_btn.setFixedWidth(160)
        apply_btn.setCursor(Qt.PointingHandCursor)
        apply_btn.clicked.connect(self._on_weekly_apply_all)
        quick_row.addWidget(apply_btn)
        quick_row.addStretch()
        wp_outer.addLayout(quick_row)

        # Separator
        sep = QFrame()
        sep.setFrameShape(QFrame.HLine)
        sep.setStyleSheet("color: #E2E8F0; max-height: 1px;")
        wp_outer.addWidget(sep)

        # ── QTabWidget: "Tất cả" + 6 carrier tabs ──
        self._sched_tab_widget = QTabWidget()
        self._sched_tab_widget.setStyleSheet("""
            QTabWidget::pane { border: 1px solid #E2E8F0; border-radius: 0 6px 6px 6px; background: #FFFFFF; }
            QTabBar::tab { padding: 6px 14px; border: 1px solid #E2E8F0; border-bottom: none; border-radius: 6px 6px 0 0; margin-right: 2px; background: #F8FAFC; }
            QTabBar::tab:selected { background: #FFFFFF; font-weight: bold; color: #2563EB; }
        """)

        day_names = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ Nhật"]
        default_times = {idx: "07:30, 13:20, 15:10, 16:10, 17:10, 18:10" if idx < 6 else "" for idx in range(7)}

        def _build_day_rows(parent_widget, cb_dict, te_dict, default_enabled=True):
            """Tạo 7 day rows cho 1 tab. Trả về layout chứa các rows."""
            day_layout = QVBoxLayout()
            day_layout.setSpacing(6)
            for idx, name in enumerate(day_names):
                day_row = QHBoxLayout()
                day_row.setSpacing(8)

                cb = QCheckBox(name)
                cb.setFixedWidth(90)
                cb.setChecked(idx < 6 and default_enabled)
                cb.setStyleSheet("font-weight: 500;")
                cb_dict[idx] = cb

                te = QLineEdit(default_times[idx] if default_enabled else "")
                te.setFixedWidth(220)
                te.setEnabled(idx < 6 and default_enabled)
                te.setPlaceholderText("VD: 07:30, 13:20, 15:10")
                te_dict[idx] = te

                cb.toggled.connect(lambda checked, i=idx, t=te: t.setEnabled(checked))

                day_row.addWidget(cb)
                day_row.addWidget(te)
                day_row.addStretch()
                day_layout.addLayout(day_row)
            day_layout.addStretch()
            return day_layout

        # ── Tab 0: "📦 Tất cả" (global schedule) ──
        global_tab = QWidget()
        global_layout = _build_day_rows(global_tab, self.weekly_day_checkboxes, self.weekly_day_time_edits, default_enabled=True)
        global_tab.setLayout(global_layout)
        self._sched_tab_widget.addTab(global_tab, "📦 Tất cả")

        # ── Tab 1-6: mỗi carrier 1 tab ──
        carrier_labels = [("GHN", "ghn"), ("J&T", "jt"), ("VietNam Post", "vnp"),
                          ("Best Express", "best"), ("Viettel Post", "viettel"), ("J&T Cargo", "jtc")]
        for c_label, c_key in carrier_labels:
            carrier_tab = QWidget()
            c_outer = QVBoxLayout(carrier_tab)
            c_outer.setContentsMargins(8, 8, 8, 8)
            c_outer.setSpacing(8)

            # Checkbox "Dùng lịch riêng"
            use_custom_cb = QCheckBox(f"🕐 Dùng lịch riêng cho {c_label}")
            use_custom_cb.setStyleSheet("font-weight: bold; color: #D97706;")
            self._carrier_use_custom_cb[c_key] = use_custom_cb
            c_outer.addWidget(use_custom_cb)

            # Day rows — disabled by default, enabled when checkbox checked
            cb_dict: dict[int, QCheckBox] = {}
            te_dict: dict[int, QLineEdit] = {}
            day_rows_layout = _build_day_rows(carrier_tab, cb_dict, te_dict, default_enabled=False)
            self._carrier_day_checkboxes[c_key] = cb_dict
            self._carrier_day_time_edits[c_key] = te_dict

            # Wrap day rows in a QWidget so we can enable/disable them
            day_rows_widget = QWidget()
            day_rows_widget.setLayout(day_rows_layout)
            day_rows_widget.setEnabled(False)
            c_outer.addWidget(day_rows_widget)

            # Connect checkbox → enable/disable day rows + copy global template
            use_custom_cb.toggled.connect(lambda checked, w=day_rows_widget, ck=c_key: self._on_carrier_custom_toggled(checked, w, ck))
            c_outer.addStretch()

            self._sched_tab_widget.addTab(carrier_tab, c_label)

        wp_outer.addWidget(self._sched_tab_widget)
        wp_outer.addStretch()
        self.weekly_panel.show()
        gb_layout.addWidget(self.weekly_panel)

        self.sched_button_group.buttonClicked.connect(self._on_schedule_mode_changed)

        layout.addWidget(gb)

        # ── Chỉ báo trạng thái lịch trình — hiển thị NGAY TRONG TAB, không phụ thuộc log/thanh dưới ──
        self.sched_status_label = QLabel("💤 Chưa bật lịch — bấm ▶ CHẠY để kích hoạt lịch trình")
        self.sched_status_label.setWordWrap(True)
        self.sched_status_label.setStyleSheet(
            "font-weight: bold; font-size: 13px; padding: 12px 16px;"
            "background: #F1F5F9; color: #64748B; border-radius: 8px;")
        layout.addWidget(self.sched_status_label)
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

        # ── Checkbox gửi Order ID lên API ──
        api_row = QHBoxLayout()
        self._test_send_api_cb = QCheckBox("📤 Gửi Order ID lên API")
        self._test_send_api_cb.setChecked(True)
        self._test_send_api_cb.setCursor(Qt.PointingHandCursor)
        api_row.addWidget(self._test_send_api_cb)
        api_row.addStretch()
        layout.addLayout(api_row)

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
        template = self._template_real
        if not master or not Path(master).exists():
            QMessageBox.critical(self, "Lỗi", "Chọn file Master (Combo) hợp lệ ở tab Tệp dữ liệu.")
            return
        if not template or not Path(template).exists():
            QMessageBox.critical(self, "Lỗi", "Chọn file Mẫu xuất hàng hợp lệ ở tab Tệp dữ liệu.")
            return
        out_dir = self.output_row.get_real_path() or str(BASE_DIR / "outputs")

        def _run():
            self._test_log.emit("info", "🧪 TEST: Bắt đầu tính toán...")
            # ── Tự nhận diện carrier từ tên file → xuất đúng "ĐVVC" trên báo cáo ──
            # Chỉ dùng khi TẤT CẢ file cùng 1 hãng; trộn nhiều hãng → để trống (gộp tất cả)
            detected = set()
            for p in pdfs:
                c = _detect_carrier_from_filename(Path(p).name)
                if c:
                    detected.add(c)
            carrier = detected.pop() if len(detected) == 1 else ''
            if carrier:
                self._test_log.emit("info", f"  🚚 Nhận diện carrier từ tên file: {carrier}")
            try:
                results = run_calculator(pdfs, out_dir, master, retail, template,
                                         lambda m, t='': self._test_log.emit(t, m),
                                         carrier=carrier,
                                         send_order_ids=self._test_send_api_cb.isChecked())
                for r in results:
                    self._test_log.emit("ok", f"  ✓ {r['rows']} SKU | Qty={r['tong_qty']} | Sold={r['tong_sold']} | Promo={r['tong_promo']}")
                    for key, lbl in [('xlsx_report', '📊')]:
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
        auto_print = self.auto_print_cb.isChecked()
        if not auto_print:
            QMessageBox.warning(self, "Cảnh báo", "Tick 'In tự động ra máy in' ở tab Cấu hình In.")
            return
        pdf_settings = self._build_pdf_settings()
        batch_size = self.batch_size_spin.value()

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

        master = self._master_real
        retail = self._retail_real
        out_dir = str(Path(all_files[0]).parent) if all_files else self.output_row.get_real_path() or str(BASE_DIR / "outputs")
        printer = self.printer_combo.currentText()
        auto_print = self.auto_print_cb.isChecked()
        pdf_settings = self._build_pdf_settings()
        batch_size = self.batch_size_spin.value()

        def _run():
            # Gom tất cả file theo carrier (nhận diện từ tên file)
            do_print = auto_print and printer
            if not do_print:
                self._test_log.emit("warn", "⚠ In bị tắt — tick 'In tự động ra máy in' ở tab Cấu hình In để in")

            # Nhóm file theo carrier: {carrier: {picking: [...], shipping: [...]}}
            # File không nhận diện được hãng (vd tên cũ chỉ có thời gian) → gom vào "Tất cả"
            by_carrier = {}
            for f in all_files:
                fname = Path(f).name
                carrier = _detect_carrier_from_filename(fname) or 'Tất cả'
                if carrier not in by_carrier:
                    by_carrier[carrier] = {'picking': [], 'shipping': []}
                if 'shipping' in fname.lower() or 'vận chuyển' in fname.lower():
                    by_carrier[carrier]['shipping'].append(f)
                else:
                    # Tất cả file còn lại (picking + file không rõ loại)
                    # đều cho vào picking để chạy qua calculator —
                    # đồng bộ với main pipeline (truyền tất cả file vào calculator)
                    by_carrier[carrier]['picking'].append(f)

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
                        results = run_calculator(groups['picking'], out_dir, master, retail, self._template_real,
                                                 lambda m, t='': self._test_log.emit(t, m),
                                                 carrier=carrier,
                                                 send_order_ids=self._test_send_api_cb.isChecked())
                        for r in results:
                            self._test_log.emit("ok", f"  ✓ {r['rows']} SKU | Qty={r['tong_qty']}")
                            fp = r['files'].get('pdf_report') or r['files'].get('xlsx_report')
                            if fp and Path(fp).exists():
                                self._add_result(fp)
                                if do_print:
                                    try:
                                        self._test_log.emit("info", f"  🖨️ In báo cáo: {Path(fp).name}")
                                        _print_file(fp, printer, pdf_settings=pdf_settings, batch_size=batch_size,
                                                    log_cb=lambda m, t='': self._test_log.emit(t, m))
                                        self._test_log.emit("ok", f"  🖨️ Báo cáo: {Path(fp).name}")
                                    except Exception as e:
                                        self._test_log.emit("err", f"  ✗ Lỗi in báo cáo: {e}")
                    except Exception as e:
                        self._test_log.emit("err", f"  ✗ Lỗi tính toán [{carrier}]: {e}")

            self._test_log.emit("bold_ok", "✅ Hoàn tất toàn bộ")

        threading.Thread(target=_run, daemon=True).start()

    # ═══════════════════════════════════════════════════════
    # TAB 5: TỔNG HỢP BÁO CÁO
    # ═══════════════════════════════════════════════════════
    def _build_aggregate_tab(self):
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)

        w = QWidget()
        w.setObjectName("scrollContent")
        layout = QVBoxLayout(w)
        layout.setContentsMargins(24, 24, 24, 24)
        layout.setSpacing(16)

        gb = QGroupBox("📊 Tổng hợp nhiều file báo cáo thành 1 file duy nhất")
        gb_layout = QVBoxLayout(gb)
        gb_layout.setSpacing(8)

        # ── Description ──
        desc = QLabel("Chọn các file Excel báo cáo (Phieu_xuat_hang_*.xlsx) để gộp lại.\n"
                      "Dữ liệu sẽ được cộng dồn theo SKU từ tất cả các file.")
        desc.setStyleSheet("color: #64748B; font-size: 12px; padding: 4px 0;")
        desc.setWordWrap(True)
        gb_layout.addWidget(desc)

        # ── File list ──
        btn_row = QHBoxLayout()
        btn_add = QPushButton("📂 Thêm file báo cáo...")
        btn_add.setObjectName("browseBtn")
        btn_add.setCursor(Qt.PointingHandCursor)
        btn_add.clicked.connect(self._aggregate_select_files)
        btn_row.addWidget(btn_add)

        btn_today = QPushButton("📅 Tổng hợp hôm nay")
        btn_today.setObjectName("schedBtn")
        btn_today.setCursor(Qt.PointingHandCursor)
        btn_today.clicked.connect(self._on_aggregate_today)
        btn_row.addWidget(btn_today)

        btn_clear = QPushButton("🗑 Xóa danh sách")
        btn_clear.setObjectName("smallBtn")
        btn_clear.setCursor(Qt.PointingHandCursor)
        btn_clear.clicked.connect(lambda: self._aggregate_file_list.clear())
        btn_row.addWidget(btn_clear)
        btn_row.addStretch()
        gb_layout.addLayout(btn_row)

        self._aggregate_file_list = QListWidget()
        self._aggregate_file_list.setMaximumHeight(150)
        self._aggregate_file_list.setObjectName("resultList")
        gb_layout.addWidget(self._aggregate_file_list)

        # ── Run button ──
        self._aggregate_run_btn = QPushButton("▶ Tổng hợp")
        self._aggregate_run_btn.setObjectName("schedBtn")
        self._aggregate_run_btn.setCursor(Qt.PointingHandCursor)
        self._aggregate_run_btn.clicked.connect(self._on_run_aggregate)
        gb_layout.addWidget(self._aggregate_run_btn)

        # ── Aggregate log ──
        self._aggregate_log = QTextEdit()
        self._aggregate_log.setObjectName("logView")
        self._aggregate_log.setReadOnly(True)
        self._aggregate_log.setMaximumHeight(250)
        gb_layout.addWidget(self._aggregate_log)

        layout.addWidget(gb)
        layout.addStretch()
        scroll.setWidget(w)
        return scroll

    def _aggregate_select_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Chọn file báo cáo Excel", "",
                                                 "Excel Files (*.xlsx)")
        for f in files:
            self._aggregate_file_list.addItem(f)

    def _on_aggregate_today(self):
        """Tìm tất cả file Phieu_xuat_hang_*.xlsx trong thư mục hôm nay và thêm vào danh sách."""
        base_dir = self.output_row.get_real_path() or str(BASE_DIR / "outputs")
        today_str = datetime.now().strftime('%Y-%m-%d')
        today_dir = Path(base_dir) / today_str

        if not today_dir.exists():
            QMessageBox.warning(self, "Thông báo",
                f"Thư mục hôm nay chưa tồn tại:\n{today_dir}")
            return

        xlsx_files = sorted(today_dir.glob("**/Phieu_xuat_hang_*.xlsx"))
        if not xlsx_files:
            QMessageBox.information(self, "Thông báo",
                f"Không tìm thấy file Phieu_xuat_hang_*.xlsx nào trong:\n{today_dir}")
            return

        # Xóa danh sách cũ và thêm file hôm nay vào
        self._aggregate_file_list.clear()
        for f in xlsx_files:
            self._aggregate_file_list.addItem(str(f))

        self._ag_log("info", f"📅 Đã tìm thấy {len(xlsx_files)} file báo cáo hôm nay ({today_str})")
        for f in xlsx_files:
            self._ag_log("dim", f"  📄 {f.name}")

        # Tự động chạy tổng hợp luôn
        self._on_run_aggregate()

    def _on_run_aggregate(self):
        files = [self._aggregate_file_list.item(i).text()
                 for i in range(self._aggregate_file_list.count())]
        if not files:
            QMessageBox.warning(self, "Cảnh báo", "Chọn ít nhất 1 file báo cáo Excel.")
            return
        template = self._template_real
        if not template or not Path(template).exists():
            QMessageBox.critical(self, "Lỗi", "Chọn file Mẫu xuất hàng hợp lệ ở tab Tệp dữ liệu.")
            return
        out_dir = self.output_row.get_real_path() or str(BASE_DIR / "outputs")

        self._aggregate_log.clear()
        self._aggregate_run_btn.setEnabled(False)

        def _run():
            self._ag_log("info", f"📊 Bắt đầu tổng hợp {len(files)} file...")
            for f in files:
                self._ag_log("dim", f"  📄 {Path(f).name}")
            try:
                from calculator import aggregate_reports  # Lazy import — chỉ load khi tổng hợp
                result = aggregate_reports(files, out_dir, template)
                self._ag_log("ok", f"  ✓ {result['rows']} SKU | Qty={result['tong_qty']} | "
                                    f"Sold={result['tong_sold']} | Promo={result['tong_promo']}")
                fp = result['files'].get('xlsx_report')
                if fp and Path(fp).exists():
                    self._add_result(fp)
                    self._ag_log("ok", f"  📊 {Path(fp).name}")
                pdf_fp = result['files'].get('pdf_report')
                if pdf_fp and Path(pdf_fp).exists():
                    self._add_result(pdf_fp, label='📄')
                self._ag_log("bold_ok", "✅ Tổng hợp hoàn tất")
            except Exception as e:
                self._ag_log("err", f"✗ Lỗi: {e}")
            finally:
                self._aggregate_run_btn.setEnabled(True)

        threading.Thread(target=_run, daemon=True).start()

    def _ag_log(self, tag: str, msg: str):
        """Ghi log vào aggregate log view (thread-safe qua signal)."""
        self._test_log.emit(tag, f"[Tổng hợp] {msg}")

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

    def _update_template_status(self):
        self.template_row.update_excel_status()

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

    def _on_template_changed(self, path: str):
        self._template_real = path
        self._update_template_status()

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
    def _collect_config(self, carrier_filter=None) -> dict:
        """Thu thập cấu hình từ UI. Nếu carrier_filter được cung cấp, chỉ include carrier trong filter."""
        carriers_to_process = []
        carrier_keys = ["ghn", "jt", "vnp", "best", "viettel", "jtc"]
        carrier_names = ["GHN", "J&T", "VietNam Post", "Best Express", "Viettel Post", "J&T Cargo VN"]
        for key, name in zip(carrier_keys, carrier_names):
            if self.carrier_checkboxes[key].isChecked():
                if carrier_filter is not None and name not in carrier_filter:
                    continue  # Skip carrier không nằm trong filter
                val = self.carrier_spinboxes[key].value()
                carriers_to_process.append((name, val))
            # unchecked = bỏ qua hãng này hoàn toàn

        return {
            "cookie": self._cookie_real,
            "output_dir": self.output_row.get_real_path() or str(BASE_DIR / "outputs"),
            "master": self._master_real,
            "retail": self._retail_real,
            "template": self._template_real,
            "carriers": carriers_to_process,
            "auto_print": self.auto_print_cb.isChecked(),
            "printer": self.printer_combo.currentText(),
            "test_mode": self.test_mode_cb.isChecked(),
            "exclude_pre_orders": self.exclude_pre_orders_cb.isChecked(),
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
            # Mở thư mục con mới nhất (theo giờ) nếu có
            subdirs = sorted([p for p in d.iterdir() if p.is_dir()], reverse=True)
            if subdirs:
                os.startfile(str(subdirs[0]))
            else:
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

        # ── Tính output_dir với date + time subfolder ──
        base_dir = config['output_dir']
        now = datetime.now()
        today_str = now.strftime('%Y-%m-%d')
        time_str = now.strftime('%H-%M-%S')
        out_dir = str(Path(base_dir) / today_str / time_str)
        os.makedirs(out_dir, exist_ok=True)
        config['output_dir'] = out_dir  # Ghi đè = path có ngày + giờ

        self.running = True
        self._set_buttons("running")
        self._clear_log()
        self._log_html("bold_ok", "▶ Bắt đầu...")
        self.status_label.setText("⏳ Đang xử lý...")
        self.status_label.setStyleSheet("color: #D97706; font-weight: bold; font-size: 14px;")
        self.trigger_job.emit(config)

    @staticmethod
    def _parse_weekly_times(day_checkboxes, day_time_edits, day_names):
        """Parse time strings from day rows. Returns {day_idx: [(h, m), ...]} or None if error."""
        result = {}
        for idx in range(7):
            if not day_checkboxes[idx].isChecked():
                continue
            time_text = day_time_edits[idx].text().strip()
            if not time_text:
                continue
            parts = [t.strip() for t in time_text.split(",") if t.strip()]
            parsed = []
            for t in parts:
                try:
                    h, m = t.split(":")
                    h_int, m_int = int(h), int(m)
                    if not (0 <= h_int <= 23 and 0 <= m_int <= 59):
                        raise ValueError
                    parsed.append((h_int, m_int))
                except (ValueError, TypeError):
                    return None, f"Giờ không hợp lệ cho {day_names[idx]}: '{t}'. Nhập dạng HH:MM (0-23:0-59)."
            if parsed:
                result[idx] = sorted(parsed)
        return result, None

    def _on_run_schedule(self):
        if not Path(self._cookie_real).exists():
            QMessageBox.critical(self, "Lỗi",
                "Chưa có file cookie TikTok hợp lệ!\n\n"
                "Mở tab 'Tệp dữ liệu' → bấm 'Chọn' cạnh '🍪 Cookie (JSON)'\n"
                "và chọn file cookie đã export từ seller-vn.tiktok.com.\n"
                "Không có cookie thì bot không đăng nhập được để chạy.")
            return
        mode = self._sched_mode
        if mode == "once":
            self._on_run_now()
            return
        if mode == "weekly":
            day_names = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ Nhật"]
            self._sched_weekly_config = {}
            self._sched_carrier_custom = {}
            has_any = False

            # ── Parse global schedule ("Tất cả" tab) ──
            global_config, err = self._parse_weekly_times(self.weekly_day_checkboxes, self.weekly_day_time_edits, day_names)
            if err:
                QMessageBox.critical(self, "Lỗi", err)
                return
            self._sched_weekly_config[""] = global_config
            if global_config:
                has_any = True

            # ── Parse per-carrier schedules ──
            carrier_keys = ["ghn", "jt", "vnp", "best", "viettel", "jtc"]
            for c_key in carrier_keys:
                use_custom = self._carrier_use_custom_cb.get(c_key) and self._carrier_use_custom_cb[c_key].isChecked()
                self._sched_carrier_custom[c_key] = use_custom
                if use_custom:
                    carrier_config, err = self._parse_weekly_times(
                        self._carrier_day_checkboxes[c_key],
                        self._carrier_day_time_edits[c_key],
                        day_names)
                    if err:
                        carrier_names = {"ghn": "GHN", "jt": "J&T", "vnp": "VietNam Post", "best": "Best Express", "viettel": "Viettel Post", "jtc": "J&T Cargo"}
                        QMessageBox.critical(self, "Lỗi", f"[{carrier_names.get(c_key, c_key)}] {err}")
                        return
                    self._sched_weekly_config[c_key] = carrier_config
                    if carrier_config:
                        has_any = True

            if not has_any:
                QMessageBox.critical(self, "Lỗi",
                    "Vui lòng chọn ít nhất một ngày và nhập ít nhất một khung giờ (global hoặc carrier).")
                return
        elif mode == "interval":
            self._sched_interval_hours = self.interval_spin.value()

        self._clear_log()
        self._clear_results()
        self._log_html("info", f"⏰ Hẹn giờ: {mode}")
        if mode == "interval":
            self._log_html("info", f"   Chạy mỗi {self._sched_interval_hours} giờ")
        elif mode == "weekly":
            day_names = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ Nhật"]
            # Log global schedule
            if self._sched_weekly_config.get(""):
                self._log_html("info", "   📦 Lịch chung (Tất cả):")
                for idx in range(7):
                    if idx in self._sched_weekly_config[""]:
                        times_str = ", ".join(f"{h:02d}:{m:02d}" for h, m in self._sched_weekly_config[""][idx])
                        self._log_html("info", f"      {day_names[idx]}: {times_str}")
            # Log per-carrier schedules
            carrier_names = {"ghn": "GHN", "jt": "J&T", "vnp": "VietNam Post", "best": "Best Express", "viettel": "Viettel Post", "jtc": "J&T Cargo"}
            for c_key in ["ghn", "jt", "vnp", "best", "viettel", "jtc"]:
                if self._sched_carrier_custom.get(c_key) and self._sched_weekly_config.get(c_key):
                    self._log_html("info", f"   🚚 {carrier_names[c_key]} (lịch riêng):")
                    for idx in range(7):
                        if idx in self._sched_weekly_config[c_key]:
                            times_str = ", ".join(f"{h:02d}:{m:02d}" for h, m in self._sched_weekly_config[c_key][idx])
                            self._log_html("info", f"      {day_names[idx]}: {times_str}")

        self._sched_mode = mode
        self.scheduler_active = True
        self._sched_last_run = None
        self._set_buttons("scheduled")
        self._sched_timer.start()

        # ── Phản hồi TỨC THÌ: tính ngay giờ chạy kế tiếp (không đợi 1s timer tick) ──
        if mode == "weekly":
            self._sched_next_run = self._calc_next_weekly_run(datetime.now())
        else:  # interval
            self._sched_next_run = datetime.now() + timedelta(hours=self._sched_interval_hours)
        self._update_countdown()
        if self._sched_next_run:
            self._log_html("bold_ok", f"🟢 ĐÃ BẬT LỊCH (chế độ {mode}) — {self.sched_status_label.text()}")
        else:
            self._log_html("warn", "⚠ Không tìm thấy giờ chạy nào trong 8 ngày tới — kiểm tra lại: ngày đã tick, giờ đã nhập đúng định dạng HH:MM.")

    def _on_stop(self):
        self.scheduler_active = False
        self._sched_timer.stop()
        self.running = False
        # ── Gọi TRỰC TIẾP stop_job (threading.Event.set() là thread-safe) ──
        # KHÔNG dùng invokeMethod với QueuedConnection vì worker thread
        # đang block trong subprocess.run / Playwright wait → event queue
        # không được xử lý → stop_event không bao giờ được set!
        self._worker.stop_job()
        # ── Force-close browser để hủy tiến trình con ngay lập tức ──
        self._worker.shutdown_browser()
        self._set_buttons("idle")
        self.status_label.setText("⏹ Đã dừng")
        self.status_label.setStyleSheet("color: #DC2626; font-weight: bold; font-size: 14px;")
        self._log_html("warn", "⏹ Đã dừng hệ thống")
        if hasattr(self, 'sched_status_label'):
            self.sched_status_label.setText("⏹ Lịch đã dừng — bấm ▶ CHẠY để kích hoạt lại")
            self.sched_status_label.setStyleSheet(
                "font-weight: bold; font-size: 13px; padding: 12px 16px;"
                "background: #FEF2F2; color: #DC2626; border-radius: 8px;")

    def _on_schedule_mode_changed(self, btn: QRadioButton):
        mode = btn.property("mode")
        self._sched_mode = mode
        self.interval_panel.setVisible(mode == "interval")
        self.weekly_panel.setVisible(mode == "weekly")

    # ═══════════════════════════════════════════════════════
    # WEEKLY SCHEDULER HELPERS
    # ═══════════════════════════════════════════════════════
    def _on_weekly_apply_all(self):
        """Copy nội dung ô giờ mẫu vào tất cả các ngày đang checked (global + carrier tabs đang bật lịch riêng)."""
        master_text = self.weekly_master_time_edit.text()
        # Global tab
        for idx in range(7):
            if self.weekly_day_checkboxes[idx].isChecked():
                self.weekly_day_time_edits[idx].setText(master_text)
        # Carrier tabs — chỉ copy vào tab đang bật "Dùng lịch riêng"
        for c_key in self._carrier_day_checkboxes:
            if self._carrier_use_custom_cb.get(c_key) and self._carrier_use_custom_cb[c_key].isChecked():
                for idx in range(7):
                    if self._carrier_day_checkboxes[c_key][idx].isChecked():
                        self._carrier_day_time_edits[c_key][idx].setText(master_text)

    def _on_carrier_custom_toggled(self, checked: bool, day_rows_widget, carrier_key: str):
        """Khi bật/tắt 'Dùng lịch riêng' cho 1 carrier → enable/disable day rows + copy global template."""
        day_rows_widget.setEnabled(checked)
        if checked:
            # Copy global template vào carrier tab khi bật lịch riêng
            for idx in range(7):
                if self._carrier_day_checkboxes[carrier_key][idx].isChecked():
                    self._carrier_day_time_edits[carrier_key][idx].setText(
                        self.weekly_day_time_edits[idx].text()
                    )

    def _find_next_run_today(self, now: datetime, schedule: dict) -> datetime | None:
        """Tìm timeslot tiếp theo trong ngày hôm nay cho 1 schedule cụ thể. Trả về None nếu hôm nay hết slot."""
        day_idx = now.weekday()  # Python: 0=Monday → khớp với idx của ta
        slots = schedule.get(day_idx, [])
        best = None
        for h, m in slots:
            candidate = now.replace(hour=h, minute=m, second=0, microsecond=0)
            if candidate > now and (best is None or candidate < best):
                best = candidate
        return best

    def _calc_next_weekly_run(self, from_time: datetime) -> datetime | None:
        """Quét tối đa 8 ngày tới, tìm timeslot sớm nhất từ TẤT CẢ schedules. Dùng cho countdown display."""
        best = None
        # Gom tất cả schedules (global + per-carrier)
        all_schedules = [s for s in self._sched_weekly_config.values() if s]
        for schedule in all_schedules:
            for offset in range(8):
                check_date = from_time.date() + timedelta(days=offset)
                day_idx = check_date.weekday()
                day_slots = schedule.get(day_idx, [])
                for h, m in day_slots:
                    candidate = datetime(check_date.year, check_date.month, check_date.day, h, m, 0, 0)
                    if candidate > from_time and (best is None or candidate < best):
                        best = candidate
                # Nếu đã tìm thấy slot trong ngày đang xét từ schedule này, dừng quét schedule này
                # (vẫn tiếp tục quét schedule khác vì có thể có slot sớm hơn)
        return best

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
        elif self._sched_mode == "weekly":
            now = datetime.now()
            # ── Gom carrier nào đến giờ chạy ──
            due_carriers = []  # list of carrier names (str)
            all_carrier_keys = ["ghn", "jt", "vnp", "best", "viettel", "jtc"]
            carrier_names_map = {"ghn": "GHN", "jt": "J&T", "vnp": "VietNam Post", "best": "Best Express", "viettel": "Viettel Post", "jtc": "J&T Cargo"}

            # Tìm next slot sớm nhất để hiển thị countdown
            best_next = None

            for c_name in all_carrier_keys:
                # Chọn schedule: carrier riêng nếu có, không thì global
                if self._sched_carrier_custom.get(c_name) and c_name in self._sched_weekly_config:
                    sched = self._sched_weekly_config[c_name]
                else:
                    sched = self._sched_weekly_config.get("", {})

                if not sched:
                    continue

                next_slot = self._find_next_run_today(now, sched)
                if next_slot is not None:
                    diff = (next_slot - now).total_seconds()
                    if diff <= 1:
                        # ── Carrier này đến giờ chạy ──
                        if not self._sched_last_run or (now - self._sched_last_run).total_seconds() > 60:
                            due_carriers.append(carrier_names_map[c_name])
                    if best_next is None or next_slot < best_next:
                        best_next = next_slot

            if due_carriers:
                # Chạy job với carrier filter
                self._sched_next_run = best_next
                self._execute_scheduled_job(carrier_filter=due_carriers)
            elif best_next is not None:
                self._sched_next_run = best_next
                self._update_countdown()
            else:
                self._sched_next_run = self._calc_next_weekly_run(now)
                self._update_countdown()

    def _execute_scheduled_job(self, carrier_filter=None):
        if self.running:
            self._log_html("dim", "⏭ Bỏ qua chu kỳ — job trước vẫn đang chạy")
            return
        config = self._collect_config(carrier_filter=carrier_filter)
        if not config['carriers']:
            carrier_info = f" [{', '.join(carrier_filter)}]" if carrier_filter else ""
            self._log_html("warn", f"⏭ Bỏ qua chu kỳ{carrier_info} — không có hãng nào được chọn")
            return

        # ── Tính output_dir với date + time subfolder ──
        base_dir = config['output_dir']
        now = datetime.now()
        today_str = now.strftime('%Y-%m-%d')
        time_str = now.strftime('%H-%M-%S')
        out_dir = str(Path(base_dir) / today_str / time_str)
        os.makedirs(out_dir, exist_ok=True)
        config['output_dir'] = out_dir

        self.running = True
        self._clear_results()
        carrier_info = f" [{', '.join(carrier_filter)}]" if carrier_filter else ""
        self.status_label.setText(f"🔄 Đang chạy tác vụ tự động{carrier_info}...")
        self.status_label.setStyleSheet("color: #D97706; font-weight: bold; font-size: 14px;")
        self.trigger_job.emit(config)

    def _update_countdown(self):
        if self._sched_next_run:
            remaining = self._sched_next_run - datetime.now()
            secs = max(0, int(remaining.total_seconds()))
            h, m = secs // 3600, (secs % 3600) // 60
            d = h // 24
            hh = h % 24
            when_str = f"⏳ {d} ngày {hh}h{m:02d} nữa" if d > 0 else f"⏳ {hh}h{m:02d} nữa"
            self.status_label.setText(when_str)
            self.status_label.setStyleSheet("color: #2563EB; font-weight: bold; font-size: 14px;")

            # ── Cập nhật chỉ báo trong tab Lịch trình ──
            if hasattr(self, 'sched_status_label'):
                target = self._sched_next_run
                now = datetime.now()
                day_diff = (target.date() - now.date()).days
                if day_diff == 0:
                    day_txt = "hôm nay"
                elif day_diff == 1:
                    day_txt = "ngày mai"
                else:
                    day_names = ["Thứ 2", "Thứ 3", "Thứ 4", "Thứ 5", "Thứ 6", "Thứ 7", "Chủ Nhật"]
                    day_txt = day_names[target.weekday()]
                self.sched_status_label.setText(
                    f"🟢 LỊCH ĐÃ KÍCH HOẠT — lần chạy kế tiếp: {day_txt} lúc {target.strftime('%H:%M')}  ({when_str})")
                self.sched_status_label.setStyleSheet(
                    "font-weight: bold; font-size: 13px; padding: 12px 16px;"
                    "background: #ECFDF5; color: #047857; border-radius: 8px;")

    # ═══════════════════════════════════════════════════════
    # CONFIG PERSISTENCE
    # ═══════════════════════════════════════════════════════
    def _save_config(self):
        """Lưu cấu hình hiện tại ra file JSON."""
        carrier_keys = ["ghn", "jt", "vnp", "best", "viettel", "jtc"]

        # ── Serialize weekly schedule (global + per-carrier) ──
        weekly_data = {}
        # Global ("Tất cả" tab)
        weekly_data[""] = {
            'days': {str(idx): {
                'checked': self.weekly_day_checkboxes[idx].isChecked(),
                'times': self.weekly_day_time_edits[idx].text(),
            } for idx in range(7)},
        }
        # Per-carrier
        for c_key in carrier_keys:
            if c_key in self._carrier_day_checkboxes:
                weekly_data[c_key] = {
                    'use_custom': self._carrier_use_custom_cb.get(c_key) and self._carrier_use_custom_cb[c_key].isChecked(),
                    'days': {str(idx): {
                        'checked': self._carrier_day_checkboxes[c_key][idx].isChecked(),
                        'times': self._carrier_day_time_edits[c_key][idx].text(),
                    } for idx in range(7)},
                }

        data = {
            'cookie': self._cookie_real,
            'master': self._master_real,
            'retail': self._retail_real,
            'template': self._template_real,
            'output_dir': self.output_row.get_real_path(),
            'carriers': {k: {
                'checked': self.carrier_checkboxes[k].isChecked(),
                'count': self.carrier_spinboxes[k].value(),
            } for k in carrier_keys},
            'auto_print': self.auto_print_cb.isChecked(),
            'printer': self.printer_combo.currentText(),
            'duplex': self.duplex_combo.currentIndex(),
            'batch_size': self.batch_size_spin.value(),
            'exclude_pre_orders': self.exclude_pre_orders_cb.isChecked(),
            'sched_mode': self._sched_mode,
            'sched_interval_hours': self._sched_interval_hours,
            'weekly_schedule': weekly_data,  # NEW: per-carrier schedule
        }
        try:
            self._config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
        except OSError:
            pass

    def _load_config(self):
        """Nạp cấu hình từ file JSON (nếu có)."""
        if not self._config_path.exists():
            return
        try:
            data = json.loads(self._config_path.read_text(encoding='utf-8'))
        except (json.JSONDecodeError, OSError):
            return

        for key, attr in [('cookie', '_cookie_real'), ('master', '_master_real'),
                          ('retail', '_retail_real'), ('template', '_template_real')]:
            saved = data.get(key, '')
            if saved and Path(saved).exists():
                setattr(self, attr, saved)

        saved_out = data.get('output_dir', '')
        if saved_out and Path(saved_out).exists():
            self.output_row.set_path(saved_out)
            self.output_row._update_status_dir()

        saved_carriers = data.get('carriers', {})
        for k, v in saved_carriers.items():
            if k in self.carrier_checkboxes:
                self.carrier_checkboxes[k].setChecked(v.get('checked', True))
                self.carrier_spinboxes[k].setValue(v.get('count', 0))

        self.auto_print_cb.setChecked(data.get('auto_print', False))
        saved_printer = data.get('printer', '')
        if saved_printer:
            idx = self.printer_combo.findText(saved_printer)
            if idx >= 0: self.printer_combo.setCurrentIndex(idx)
        idx_duplex = data.get('duplex', -1)
        if 0 <= idx_duplex < self.duplex_combo.count():
            self.duplex_combo.setCurrentIndex(idx_duplex)
        self.batch_size_spin.setValue(data.get('batch_size', 0))
        self.exclude_pre_orders_cb.setChecked(data.get('exclude_pre_orders', True))

        self._sched_mode = data.get('sched_mode', 'weekly')
        self._sched_interval_hours = data.get('sched_interval_hours', 1)

        # ── Restore weekly schedule (global + per-carrier) ──
        weekly_data = data.get('weekly_schedule', {})
        if weekly_data:
            # Global ("Tất cả" tab)
            global_data = weekly_data.get("", {}).get('days', {})
            for idx_str, day_data in global_data.items():
                idx = int(idx_str)
                if idx in self.weekly_day_checkboxes:
                    self.weekly_day_checkboxes[idx].setChecked(day_data.get('checked', idx < 6))
                    self.weekly_day_time_edits[idx].setText(day_data.get('times', ''))
                    self.weekly_day_time_edits[idx].setEnabled(day_data.get('checked', idx < 6))
            # Per-carrier
            carrier_keys = ["ghn", "jt", "vnp", "best", "viettel", "jtc"]
            for c_key in carrier_keys:
                c_data = weekly_data.get(c_key, {})
                if c_key in self._carrier_use_custom_cb:
                    use_custom = c_data.get('use_custom', False)
                    self._carrier_use_custom_cb[c_key].setChecked(use_custom)
                if c_key in self._carrier_day_checkboxes:
                    days_data = c_data.get('days', {})
                    for idx_str, day_data in days_data.items():
                        idx = int(idx_str)
                        if idx in self._carrier_day_checkboxes[c_key]:
                            self._carrier_day_checkboxes[c_key][idx].setChecked(day_data.get('checked', idx < 6))
                            self._carrier_day_time_edits[c_key][idx].setText(day_data.get('times', ''))
                            self._carrier_day_time_edits[c_key][idx].setEnabled(day_data.get('checked', idx < 6) and use_custom)

        self.cookie_row.set_path(self._cookie_real)
        self.master_row.set_path(self._master_real)
        self.retail_row.set_path(self._retail_real)
        self.template_row.set_path(self._template_real)
        self._update_cookie_status()
        self._update_master_status()
        self._update_retail_status()
        self._update_template_status()

    # ═══════════════════════════════════════════════════════
    # CLOSE EVENT
    # ═══════════════════════════════════════════════════════
    def closeEvent(self, event):
        self.scheduler_active = False
        self._sched_timer.stop()
        self.running = False
        self._save_config()
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