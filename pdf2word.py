import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import os
import threading
from datetime import datetime
import sys
import webbrowser
import pyperclip
import qrcode
from PIL import Image, ImageTk
import win32com.client
import requests
from packaging import version

VERSION = "1.0.0"

LANGUAGES = {
    "en": {
        "title": "PDF to Word Converter",
        "select": "Select PDF File",
        "select_btn": "Browse",
        "no_file": "No file selected",
        "loaded": "File loaded successfully",
        "output": "Output File Name (Optional)",
        "output_hint": "Default: Converted_Time",
        "save": "Save Location",
        "save_btn": "Browse",
        "convert": "Convert to Word",
        "please_select": "Please select a file first!",
        "converting": "Converting with Microsoft Word...",
        "success": "Conversion completed!",
        "error": "Conversion failed",
        "word_not_found": "Microsoft Word not installed!",
        "about": "About",
        "contact": "Contact",
        "donate": "Donate",
        "donate_title": "Support the Developer",
        "update": "Update",
        "language": "Language",
        "menu": "Menu",
        "info1": "Original quality using Microsoft Word",
        "info2": "Preserves formatting, tables, images",
        "info3": "Requires Microsoft Word installed",
        "footer": "Made with ♥ by FHASH_CODE",
        "github": "GitHub",
        "github_url": "github.com/fhash-code",
        "email": "Email",
        "email_addr": "fhash_code@proton.me",
        "copy": "Copy",
        "copied": "Copied!",
        "wallet_bep20": "BEP-20 Wallet:",
        "wallet_trc20": "TRC20 Wallet (USDT):",
        "wallet_addr_bep20": "0x4d2BEF479014BbF398a4005b5D07f82752567C61",
        "wallet_addr_trc20": "TWDFUXRVCRXvEswftqFHwnkt3BXk28PYXN",
        "close": "Close",
        "scan_qr": "Scan QR to Donate",
        "support_text": "Supported: BEP-20 (Binance) and TRC20 (USDT)",
        "version": "Version",
        "update_status": "Up to date",
        "new_version": "New version available!",
        "latest": "You have the latest version",
        "checking": "Checking for updates..."
    },
    "fa": {
        "title": "تبدیل PDF به Word",
        "select": "انتخاب فایل PDF",
        "select_btn": "انتخاب",
        "no_file": "هیچ فایلی انتخاب نشده",
        "loaded": "فایل با موفقیت بارگذاری شد",
        "output": "نام فایل خروجی (اختیاری)",
        "output_hint": "پیش‌فرض: Converted_ساعت",
        "save": "مسیر ذخیره",
        "save_btn": "انتخاب",
        "convert": "تبدیل به Word",
        "please_select": "لطفاً فایل را انتخاب کنید!",
        "converting": "در حال تبدیل با Microsoft Word...",
        "success": "تبدیل با موفقیت انجام شد!",
        "error": "خطا در تبدیل",
        "word_not_found": "Microsoft Word نصب نیست!",
        "about": "درباره",
        "contact": "ارتباط",
        "donate": "حمایت",
        "donate_title": "حمایت مالی",
        "update": "بروزرسانی",
        "language": "زبان",
        "menu": "منو",
        "info1": "کیفیت اصلی با Microsoft Word",
        "info2": "حفظ فرمت، جداول و تصاویر",
        "info3": "نیاز به Microsoft Word نصب شده",
        "footer": "ساخته شده با ♥ توسط FHASH_CODE",
        "github": "گیت‌هاب",
        "github_url": "github.com/fhash-code",
        "email": "ایمیل",
        "email_addr": "fhash_code@proton.me",
        "copy": "کپی",
        "copied": "کپی شد!",
        "wallet_bep20": "کیف پول (BEP-20):",
        "wallet_trc20": "کیف پول (TRC20 - USDT):",
        "wallet_addr_bep20": "0x4d2BEF479014BbF398a4005b5D07f82752567C61",
        "wallet_addr_trc20": "TWDFUXRVCRXvEswftqFHwnkt3BXk28PYXN",
        "close": "بستن",
        "scan_qr": "اسکن QR برای حمایت",
        "support_text": "شبکه‌ها: BEP-20 (بایننس) و TRC20 (USDT)",
        "version": "نسخه",
        "update_status": "بروز است",
        "new_version": "نسخه جدید موجود است!",
        "latest": "شما آخرین نسخه را دارید",
        "checking": "در حال بررسی..."
    },
    "ar": {
        "title": "تحويل PDF إلى Word",
        "select": "اختيار ملف PDF",
        "select_btn": "تصفح",
        "no_file": "لم يتم اختيار ملف",
        "loaded": "تم تحميل الملف بنجاح",
        "output": "اسم ملف الإخراج (اختياري)",
        "output_hint": "الافتراضي: Converted_الوقت",
        "save": "موقع الحفظ",
        "save_btn": "تصفح",
        "convert": "تحويل إلى Word",
        "please_select": "الرجاء اختيار ملف أولاً!",
        "converting": "جاري التحويل باستخدام Microsoft Word...",
        "success": "تم التحويل بنجاح!",
        "error": "فشل التحويل",
        "word_not_found": "Microsoft Word غير مثبت!",
        "about": "حول",
        "contact": "اتصل بي",
        "donate": "تبرع",
        "donate_title": "دعم المطور",
        "update": "تحديث",
        "language": "اللغة",
        "menu": "القائمة",
        "info1": "جودة أصلية باستخدام Microsoft Word",
        "info2": "الحفاظ على التنسيق والجداول والصور",
        "info3": "يتطلب تثبيت Microsoft Word",
        "footer": "صنع بـ ♥ بواسطة FHASH_CODE",
        "github": "جيت هاب",
        "github_url": "github.com/fhash-code",
        "email": "البريد الإلكتروني",
        "email_addr": "fhash_code@proton.me",
        "copy": "نسخ",
        "copied": "تم النسخ!",
        "wallet_bep20": "محفظة (BEP-20):",
        "wallet_trc20": "محفظة (TRC20 - USDT):",
        "wallet_addr_bep20": "0x4d2BEF479014BbF398a4005b5D07f82752567C61",
        "wallet_addr_trc20": "TWDFUXRVCRXvEswftqFHwnkt3BXk28PYXN",
        "close": "إغلاق",
        "scan_qr": "امسح QR للتبرع",
        "support_text": "الشبكات المدعومة: BEP-20 و TRC20",
        "version": "الإصدار",
        "update_status": "أحدث إصدار",
        "new_version": "يتوفر إصدار جديد!",
        "latest": "لديك أحدث إصدار",
        "checking": "جاري التحقق من التحديثات..."
    },
    "zh": {
        "title": "PDF转Word转换器",
        "select": "选择PDF文件",
        "select_btn": "浏览",
        "no_file": "未选择文件",
        "loaded": "文件加载成功",
        "output": "输出文件名（可选）",
        "output_hint": "默认：Converted_时间",
        "save": "保存位置",
        "save_btn": "浏览",
        "convert": "转换为Word",
        "please_select": "请先选择文件！",
        "converting": "正在使用Microsoft Word转换...",
        "success": "转换成功！",
        "error": "转换失败",
        "word_not_found": "未安装Microsoft Word！",
        "about": "关于",
        "contact": "联系我",
        "donate": "捐赠",
        "donate_title": "支持开发者",
        "update": "更新",
        "language": "语言",
        "menu": "菜单",
        "info1": "使用Microsoft Word保持原始质量",
        "info2": "保留格式、表格和图片",
        "info3": "需要安装Microsoft Word",
        "footer": "用♥制作 by FHASH_CODE",
        "github": "GitHub",
        "github_url": "github.com/fhash-code",
        "email": "邮箱",
        "email_addr": "fhash_code@proton.me",
        "copy": "复制",
        "copied": "已复制！",
        "wallet_bep20": "钱包地址 (BEP-20):",
        "wallet_trc20": "钱包地址 (TRC20 - USDT):",
        "wallet_addr_bep20": "0x4d2BEF479014BbF398a4005b5D07f82752567C61",
        "wallet_addr_trc20": "TWDFUXRVCRXvEswftqFHwnkt3BXk28PYXN",
        "close": "关闭",
        "scan_qr": "扫描二维码捐赠",
        "support_text": "支持网络：BEP-20 和 TRC20",
        "version": "版本",
        "update_status": "已是最新",
        "new_version": "有新版本可用！",
        "latest": "您使用的是最新版本",
        "checking": "正在检查更新..."
    },
    "ru": {
        "title": "Конвертер PDF в Word",
        "select": "Выбрать PDF файл",
        "select_btn": "Обзор",
        "no_file": "Файл не выбран",
        "loaded": "Файл загружен",
        "output": "Имя выходного файла (опционально)",
        "output_hint": "По умолчанию: Converted_Время",
        "save": "Место сохранения",
        "save_btn": "Обзор",
        "convert": "Конвертировать в Word",
        "please_select": "Пожалуйста, сначала выберите файл!",
        "converting": "Конвертация с Microsoft Word...",
        "success": "Конвертация успешна!",
        "error": "Ошибка конвертации",
        "word_not_found": "Microsoft Word не установлен!",
        "about": "О программе",
        "contact": "Связаться",
        "donate": "Поддержать",
        "donate_title": "Поддержать разработчика",
        "update": "Обновить",
        "language": "Язык",
        "menu": "Меню",
        "info1": "Оригинальное качество через Microsoft Word",
        "info2": "Сохраняет форматирование, таблицы, изображения",
        "info3": "Требуется установленный Microsoft Word",
        "footer": "Сделано с ♥ by FHASH_CODE",
        "github": "GitHub",
        "github_url": "github.com/fhash-code",
        "email": "Email",
        "email_addr": "fhash_code@proton.me",
        "copy": "Копировать",
        "copied": "Скопировано!",
        "wallet_bep20": "Кошелек (BEP-20):",
        "wallet_trc20": "Кошелек (TRC20 - USDT):",
        "wallet_addr_bep20": "0x4d2BEF479014BbF398a4005b5D07f82752567C61",
        "wallet_addr_trc20": "TWDFUXRVCRXvEswftqFHwnkt3BXk28PYXN",
        "close": "Закрыть",
        "scan_qr": "Сканируйте QR для поддержки",
        "support_text": "Поддерживаемые сети: BEP-20 и TRC20",
        "version": "Версия",
        "update_status": "Последняя версия",
        "new_version": "Доступна новая версия!",
        "latest": "У вас последняя версия",
        "checking": "Проверка обновлений..."
    }
}


class PdfToWordApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PDF to Word Converter")
        self.root.geometry("950x750")
        self.root.minsize(800, 650)
        self.root.configure(bg='#f0f0f0')

        self.selected_file = None
        self.output_dir = os.path.expanduser("~/Documents")
        self.output_name = tk.StringVar()
        self.current_lang = "en"

        self._setup_ui()
        self._update_ui()

    def _(self, key: str) -> str:
        return LANGUAGES[self.current_lang].get(key, key)

    def _switch_language(self, lang_code: str) -> None:
        self.current_lang = lang_code
        self._update_ui()

    def _copy_to_clipboard(self, text: str, status_label: tk.Label) -> None:
        pyperclip.copy(text)
        status_label.config(text=self._("copied"), fg='green')
        self.root.after(2000, lambda: status_label.config(text=""))

    def _open_url(self, url: str) -> None:
        webbrowser.open(f"https://{url}")

    def _open_email(self, email: str) -> None:
        webbrowser.open(f"mailto:{email}")

    def _create_qr(self, data: str, size: int = 80) -> ImageTk.PhotoImage:
        qr = qrcode.QRCode(version=1, box_size=3, border=2)
        qr.add_data(data)
        qr.make(fit=True)
        img = qr.make_image(fill_color="#e74c3c", back_color="white")
        img = img.resize((size, size), Image.Resampling.LANCZOS)
        return ImageTk.PhotoImage(img)

    def _show_donate_window(self) -> None:
        win = tk.Toplevel(self.root)
        win.title(self._("donate_title"))
        win.geometry("550x600")
        win.configure(bg='#f0f0f0')
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text=self._("donate_title"), font=("Segoe UI", 14, "bold"),
                bg='#2c3e50', fg='white', pady=10).pack(fill='x')

        main_frame = tk.Frame(win, bg='#f0f0f0', padx=20, pady=15)
        main_frame.pack(fill='both', expand=True)

        # BEP-20 section with QR and copy button
        bep20_frame = tk.LabelFrame(main_frame, text="BEP-20 (Binance Smart Chain)",
                                    font=("Segoe UI", 10, "bold"), bg='#f0f0f0', padx=10, pady=10)
        bep20_frame.pack(fill='x', pady=8)

        bep20_addr = self._("wallet_addr_bep20")
        qr_img = self._create_qr(bep20_addr, 80)
        qr_label = tk.Label(bep20_frame, image=qr_img, bg='#f0f0f0')
        qr_label.image = qr_img
        qr_label.pack(pady=5)

        bep20_addr_frame = tk.Frame(bep20_frame, bg='#f0f0f0')
        bep20_addr_frame.pack(fill='x', pady=5)

        bep20_addr_label = tk.Label(bep20_addr_frame, text=bep20_addr, font=("Segoe UI", 9),
                                    bg='#f0f0f0', fg='#3498db', cursor="hand2")
        bep20_addr_label.pack(side='left', padx=5)
        bep20_addr_label.bind("<Button-1>", lambda e: self._copy_to_clipboard(bep20_addr, bep20_status))

        bep20_copy_btn = tk.Button(bep20_addr_frame, text=self._("copy"),
                                   command=lambda: self._copy_to_clipboard(bep20_addr, bep20_status),
                                   bg='#3498db', fg='white', padx=10, pady=2, relief='flat')
        bep20_copy_btn.pack(side='right', padx=5)

        bep20_status = tk.Label(bep20_frame, text="", bg='#f0f0f0', fg='green')
        bep20_status.pack()

        # TRC20 section with QR and copy button
        trc20_frame = tk.LabelFrame(main_frame, text="TRC20 (USDT - Tron)",
                                    font=("Segoe UI", 10, "bold"), bg='#f0f0f0', padx=10, pady=10)
        trc20_frame.pack(fill='x', pady=8)

        trc20_addr = self._("wallet_addr_trc20")
        trc20_qr = self._create_qr(trc20_addr, 80)
        trc20_label = tk.Label(trc20_frame, image=trc20_qr, bg='#f0f0f0')
        trc20_label.image = trc20_qr
        trc20_label.pack(pady=5)

        trc20_addr_frame = tk.Frame(trc20_frame, bg='#f0f0f0')
        trc20_addr_frame.pack(fill='x', pady=5)

        trc20_addr_label = tk.Label(trc20_addr_frame, text=trc20_addr, font=("Segoe UI", 9),
                                    bg='#f0f0f0', fg='#3498db', cursor="hand2")
        trc20_addr_label.pack(side='left', padx=5)
        trc20_addr_label.bind("<Button-1>", lambda e: self._copy_to_clipboard(trc20_addr, trc20_status))

        trc20_copy_btn = tk.Button(trc20_addr_frame, text=self._("copy"),
                                   command=lambda: self._copy_to_clipboard(trc20_addr, trc20_status),
                                   bg='#3498db', fg='white', padx=10, pady=2, relief='flat')
        trc20_copy_btn.pack(side='right', padx=5)

        trc20_status = tk.Label(trc20_frame, text="", bg='#f0f0f0', fg='green')
        trc20_status.pack()

        tk.Label(main_frame, text=self._("support_text"), font=("Segoe UI", 9),
                bg='#f0f0f0', fg='gray').pack(pady=10)

        tk.Button(main_frame, text=self._("close"), command=win.destroy,
                  bg='#e74c3c', fg='white', padx=20, pady=5, relief='flat').pack(pady=10)

    def _update_ui(self) -> None:
        self.root.title(self._("title"))
        if hasattr(self, 'title_label'):
            self.title_label.config(text=self._("title"))
        self._rebuild_menus()
        self._update_frame_texts()
        self.root.update()

    def _update_frame_texts(self) -> None:
        if hasattr(self, 'select_frame'):
            self.select_frame.config(text=self._("select"))
        if hasattr(self, 'name_frame'):
            self.name_frame.config(text=self._("output"))
        if hasattr(self, 'save_frame'):
            self.save_frame.config(text=self._("save"))
        if hasattr(self, 'select_btn'):
            self.select_btn.config(text=self._("select_btn"))
        if hasattr(self, 'save_btn'):
            self.save_btn.config(text=self._("save_btn"))
        if hasattr(self, 'convert_btn'):
            self.convert_btn.config(text=self._("convert"))
        if hasattr(self, 'donate_btn'):
            self.donate_btn.config(text=self._("donate"))
        if hasattr(self, 'file_label') and not self.selected_file:
            self.file_label.config(text=self._("no_file"))
        if hasattr(self, 'output_hint'):
            self.output_hint.config(text=self._("output_hint"))
        if hasattr(self, 'info1'):
            self.info1.config(text=self._("info1"))
            self.info2.config(text=self._("info2"))
            self.info3.config(text=self._("info3"))
        if hasattr(self, 'footer_label'):
            self.footer_label.config(text=self._("footer"))

    def _rebuild_menus(self) -> None:
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        main_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=f"☰ {self._('menu')}", menu=main_menu)
        main_menu.add_command(label=self._("about"), command=self._show_about)
        main_menu.add_separator()
        main_menu.add_command(label=self._("contact"), command=self._show_contact)
        main_menu.add_separator()
        main_menu.add_command(label=self._("update"), command=self._check_updates)

        lang_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label=f"🌐 {self._('language')}", menu=lang_menu)
        for code, name in [("fa", "فارسی"), ("en", "English"), ("ar", "العربية"), ("zh", "中文"), ("ru", "Русский")]:
            lang_menu.add_command(label=name, command=lambda c=code: self._switch_language(c))

    def _show_contact(self) -> None:
        win = tk.Toplevel(self.root)
        win.title(self._("contact"))
        win.geometry("500x280")
        win.configure(bg='#f0f0f0')
        win.resizable(False, False)
        win.transient(self.root)
        win.grab_set()

        tk.Label(win, text=self._("contact"), font=("Segoe UI", 14, "bold"),
                bg='#2c3e50', fg='white', pady=10).pack(fill='x')

        frame = tk.Frame(win, bg='#f0f0f0', padx=20, pady=20)
        frame.pack(fill='both', expand=True)

        # GitHub section with copy button
        github_frame = tk.Frame(frame, bg='#f0f0f0')
        github_frame.pack(fill='x', pady=8)

        tk.Label(github_frame, text=f"{self._('github')}: ", font=("Segoe UI", 10, "bold"),
                bg='#f0f0f0').pack(side='left')

        github_link = tk.Label(github_frame, text=self._("github_url"), font=("Segoe UI", 10, "underline"),
                               bg='#f0f0f0', fg='#3498db', cursor="hand2")
        github_link.pack(side='left', padx=5)
        github_link.bind("<Button-1>", lambda e: self._open_url(self._("github_url")))

        github_copy_btn = tk.Button(github_frame, text=self._("copy"),
                                    command=lambda: self._copy_to_clipboard(self._("github_url"), github_status),
                                    bg='#3498db', fg='white', padx=10, pady=2, relief='flat')
        github_copy_btn.pack(side='right', padx=5)

        github_status = tk.Label(frame, text="", bg='#f0f0f0', fg='green')
        github_status.pack()

        # Email section with copy button
        email_frame = tk.Frame(frame, bg='#f0f0f0')
        email_frame.pack(fill='x', pady=8)

        tk.Label(email_frame, text=f"{self._('email')}: ", font=("Segoe UI", 10, "bold"),
                bg='#f0f0f0').pack(side='left')

        email_link = tk.Label(email_frame, text=self._("email_addr"), font=("Segoe UI", 10, "underline"),
                              bg='#f0f0f0', fg='#3498db', cursor="hand2")
        email_link.pack(side='left', padx=5)
        email_link.bind("<Button-1>", lambda e: self._open_email(self._("email_addr")))

        email_copy_btn = tk.Button(email_frame, text=self._("copy"),
                                   command=lambda: self._copy_to_clipboard(self._("email_addr"), email_status),
                                   bg='#3498db', fg='white', padx=10, pady=2, relief='flat')
        email_copy_btn.pack(side='right', padx=5)

        email_status = tk.Label(frame, text="", bg='#f0f0f0', fg='green')
        email_status.pack()

        tk.Button(frame, text=self._("close"), command=win.destroy,
                  bg='#e74c3c', fg='white', padx=20, pady=5, relief='flat').pack(pady=20)

    def _check_updates(self) -> None:
        self.status.config(text=self._("checking"), fg='orange')
        self.root.update()

        def check():
            try:
                url = "https://raw.githubusercontent.com/fhash-code/pdf-to-word-converter/main/version.json"
                resp = requests.get(url, timeout=5)
                if resp.status_code == 200:
                    data = resp.json()
                    latest = data.get("version", "0")
                    if version.parse(latest) > version.parse(VERSION):
                        notes = data.get("release_notes", "")
                        dl_url = data.get("download_url", "https://github.com/fhash-code/pdf-to-word-converter/releases/latest")
                        self.root.after(0, lambda: self._show_update_prompt(latest, notes, dl_url))
                        return
                self.root.after(0, lambda: self.status.config(text=self._("latest"), fg='green'))
            except Exception:
                self.root.after(0, lambda: self.status.config(text=self._("latest"), fg='green'))

        threading.Thread(target=check, daemon=True).start()

    def _show_update_prompt(self, new_version: str, notes: str, download_url: str) -> None:
        if messagebox.askyesno(self._("update"), f"{self._('new_version')}\n\nVersion {new_version}\n\n{notes[:300]}\n\nDownload now?"):
            webbrowser.open(download_url)

    def _show_about(self) -> None:
        messagebox.showinfo(self._("about"), f"{self._('title')}\nVersion {VERSION}\n\n{self._('info1')}\n{self._('info2')}\n{self._('info3')}\n\n{self._('footer')}")

    def _setup_ui(self) -> None:
        # Configure grid weights for responsive layout
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)

        main_container = tk.Frame(self.root, bg='#f0f0f0')
        main_container.grid(row=0, column=0, sticky="nsew")
        main_container.grid_rowconfigure(1, weight=1)
        main_container.grid_columnconfigure(0, weight=1)

        # Title
        title_frame = tk.Frame(main_container, bg='#2c3e50', height=80)
        title_frame.grid(row=0, column=0, sticky="ew")
        title_frame.grid_propagate(False)
        self.title_label = tk.Label(title_frame, text=self._("title"), font=("Segoe UI", 15, "bold"),
                                    bg='#2c3e50', fg='white')
        self.title_label.pack(expand=True)

        # Main content area
        content_frame = tk.Frame(main_container, bg='#f0f0f0', padx=30, pady=20)
        content_frame.grid(row=1, column=0, sticky="nsew")
        content_frame.grid_rowconfigure(6, weight=1)
        content_frame.grid_columnconfigure(0, weight=1)

        # File selection
        self.select_frame = tk.LabelFrame(content_frame, text=self._("select"), font=("Segoe UI", 11, "bold"),
                                          bg='#f0f0f0', padx=15, pady=15)
        self.select_frame.grid(row=0, column=0, sticky="ew", pady=8)
        self.select_frame.grid_columnconfigure(0, weight=1)

        self.file_label = tk.Label(self.select_frame, text=self._("no_file"), bg='#f0f0f0', fg='gray')
        self.file_label.grid(row=0, column=0, sticky="w", padx=5)

        self.select_btn = tk.Button(self.select_frame, text=self._("select_btn"), command=self._select_file,
                                    bg='#3498db', fg='white', padx=20, pady=5, relief='flat')
        self.select_btn.grid(row=0, column=1, padx=10)

        # Output name
        self.name_frame = tk.LabelFrame(content_frame, text=self._("output"), font=("Segoe UI", 11, "bold"),
                                        bg='#f0f0f0', padx=15, pady=15)
        self.name_frame.grid(row=1, column=0, sticky="ew", pady=8)
        self.name_frame.grid_columnconfigure(0, weight=1)

        self.output_entry = tk.Entry(self.name_frame, textvariable=self.output_name, font=("Segoe UI", 11))
        self.output_entry.grid(row=0, column=0, sticky="ew", padx=10, pady=5)

        self.output_hint = tk.Label(self.name_frame, text=self._("output_hint"), font=("Segoe UI", 9),
                                    fg='gray', bg='#f0f0f0')
        self.output_hint.grid(row=1, column=0)

        # Save location
        self.save_frame = tk.LabelFrame(content_frame, text=self._("save"), font=("Segoe UI", 11, "bold"),
                                        bg='#f0f0f0', padx=15, pady=15)
        self.save_frame.grid(row=2, column=0, sticky="ew", pady=8)
        self.save_frame.grid_columnconfigure(0, weight=1)

        self.save_label = tk.Label(self.save_frame, text=self.output_dir, bg='#f0f0f0')
        self.save_label.grid(row=0, column=0, sticky="w", padx=10)

        self.save_btn = tk.Button(self.save_frame, text=self._("save_btn"), command=self._select_output_dir,
                                  bg='#2ecc71', fg='white', padx=20, pady=5, relief='flat')
        self.save_btn.grid(row=0, column=1, padx=10)

        # Buttons row
        btn_frame = tk.Frame(content_frame, bg='#f0f0f0')
        btn_frame.grid(row=3, column=0, pady=15)
        btn_frame.grid_columnconfigure(0, weight=1)
        btn_frame.grid_columnconfigure(1, weight=1)

        self.convert_btn = tk.Button(btn_frame, text=self._("convert"), command=self._convert,
                                     bg='#e74c3c', fg='white', font=("Segoe UI", 12, "bold"),
                                     padx=35, pady=8, relief='flat', state='disabled')
        self.convert_btn.grid(row=0, column=0, padx=10)

        self.donate_btn = tk.Button(btn_frame, text=self._("donate"), command=self._show_donate_window,
                                    bg='#f39c12', fg='white', font=("Segoe UI", 12, "bold"),
                                    padx=25, pady=8, relief='flat')
        self.donate_btn.grid(row=0, column=1, padx=10)

        # Status
        self.status = tk.Label(content_frame, text="", bg='#f0f0f0', fg='#e74c3c')
        self.status.grid(row=4, column=0, pady=5)

        # Progress
        self.progress = ttk.Progressbar(content_frame, mode='indeterminate', length=400)
        self.progress.grid(row=5, column=0, pady=10)

        # Info
        info_frame = tk.Frame(content_frame, bg='#f0f0f0')
        info_frame.grid(row=6, column=0, pady=10)
        self.info1 = tk.Label(info_frame, text=self._("info1"), bg='#f0f0f0', fg='gray')
        self.info1.pack()
        self.info2 = tk.Label(info_frame, text=self._("info2"), bg='#f0f0f0', fg='gray')
        self.info2.pack()
        self.info3 = tk.Label(info_frame, text=self._("info3"), bg='#f0f0f0', fg='gray')
        self.info3.pack()

        # Footer
        footer_frame = tk.Frame(main_container, bg='#2c3e50', height=35)
        footer_frame.grid(row=2, column=0, sticky="ew")
        footer_frame.grid_propagate(False)
        self.footer_label = tk.Label(footer_frame, text=self._("footer"), font=("Segoe UI", 9, "italic"),
                                     bg='#2c3e50', fg='#e74c3c')
        self.footer_label.pack(expand=True)

        self._rebuild_menus()

    def _select_file(self) -> None:
        path = filedialog.askopenfilename(title=self._("select"), filetypes=[("PDF files", "*.pdf")])
        if not path:
            return
        if not os.path.exists(path) or not path.lower().endswith('.pdf'):
            self.status.config(text="Invalid file", fg='red')
            return
        self.selected_file = path
        self.file_label.config(text=f"✅ {os.path.basename(path)} ({os.path.getsize(path) / 1024:.1f} KB)", fg='green')
        self.convert_btn.config(state='normal')
        self.status.config(text=self._("loaded"), fg='green')

    def _select_output_dir(self) -> None:
        path = filedialog.askdirectory(title=self._("save"))
        if path and os.access(path, os.W_OK):
            self.output_dir = path
            self.save_label.config(text=path)

    def _convert(self) -> None:
        if not self.selected_file:
            self.status.config(text=self._("please_select"), fg='red')
            return
        base = self.output_name.get().strip()
        if not base:
            base = f"Converted_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        else:
            base = "".join(c for c in base if c not in '<>:"/\\|?*')
        output = os.path.join(self.output_dir, f"{base}.docx")

        self.convert_btn.config(state='disabled')
        self.progress.start()
        self.status.config(text=self._("converting"), fg='orange')
        self.root.update()
        threading.Thread(target=self._convert_thread, args=(output,), daemon=True).start()

    def _convert_thread(self, output: str) -> None:
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            word.DisplayAlerts = False
            doc = word.Documents.Open(os.path.abspath(self.selected_file))
            doc.SaveAs2(output, FileFormat=16)
            doc.Close()
            word.Quit()
            self.root.after(0, lambda: self._convert_success(output))
        except Exception as e:
            self.root.after(0, lambda: self._convert_error(str(e)))

    def _convert_success(self, output: str) -> None:
        self.progress.stop()
        self.status.config(text=self._("success"), fg='green')
        if messagebox.askyesno(self._("success"), f"{self._('success')}\n\nOpen file?"):
            os.startfile(output)
        self.convert_btn.config(state='disabled')

    def _convert_error(self, error_msg: str) -> None:
        self.progress.stop()
        self.status.config(text=f"{self._('error')}: {error_msg}", fg='red')
        self.convert_btn.config(state='normal')

    def run(self) -> None:
        self.root.mainloop()


if __name__ == "__main__":
    app = PdfToWordApp()
    app.run()