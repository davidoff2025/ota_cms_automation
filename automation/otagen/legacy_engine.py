import os
import re
import sys
import json
import queue
import threading
import traceback
import subprocess
import shutil
from urllib.parse import urlparse, unquote
from datetime import datetime
from ftplib import FTP

# Tkinter is only needed for the legacy desktop GUI. In the v4 web app we import
# this module purely for its business logic and run it headless (no display), so
# the GUI imports are optional. Any attempt to use the GUI methods without a
# display will still fail loudly, but importing the pure logic always works.
try:  # pragma: no cover - environment dependent
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk
    _TK_AVAILABLE = True
except Exception:  # pragma: no cover
    tk = None
    filedialog = messagebox = scrolledtext = simpledialog = ttk = None
    _TK_AVAILABLE = False

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


AWS_PREFIX = os.getenv("OTA_AWS_PREFIX", "https://ota-dl.vidaahub.com")
# Credentials must come from the environment — never hard-code secrets in source.
# (v4's active FTP path is app/core/transfer.py, which reads app/core/config.py.)
FTP_USER = os.getenv("OTA_FTP_USER", "ftpuser")
FTP_PASS = os.getenv("OTA_FTP_PASS", "")
OTA_MESSAGE_TEMPLATES_FILE = "ota_message_templates.json"
DEVICE_HISTORY_SEED_FILE = "device_history_seed.json"

CONFIG_FILE = os.path.join(
    os.path.expanduser("~"),
    ".ota_deploy_tool_config.json"
)
USER_OTA_MESSAGE_TEMPLATES_FILE = os.path.join(
    os.path.expanduser("~"),
    ".ota_deploy_tool_message_templates.json"
)

OUTPUT_HEADERS = [
    "设备类型代码",
    "设备类型名称",
    "特征码",
    "内部机型信息",
    "设备扩展信息",
    "OTA类型",
    "源版本",
    "目标完整版本",
    "目标版本",
    "品牌组",
    "区域组",
    "MAC组",
    "定向组",
    "升级文件地址",
    "升级文件的MD5值",
    "文件大小",
    "SHA256",
    "EULA文件地址",
]

COLUMN_ALIASES = {
    "机型信息": ["机型", "机型信息", "内部机型", "内部机型信息", "hisense/songshan", "model", "model name"],
    "机器扩展码": [
        "机器扩展码", "扩展码", "扩展信息", "机器扩展信息", "设备扩展信息",
        "extension info", "extension code",
    ],
    "OTA类型": ["OTA类型", "OTA升级类型", "升级类型", "ota type", "upgrade type"],
    "源版本": ["源版本", "source version", "源文件版本"],
    "目标版本": ["目标版本", "target version", "完整版本", "目标文件版本"],
    "版本变化": ["版本变化", "version change", "版本差异"],
    "升级包地址": ["升级包地址", "整包地址", "整包文件地址", "完整包地址", "全包地址", "OTA整包地址", "ZIP地址", "ZIP文件地址", "升级文件地址", "ftp path", "ftp url", "upgrade url", "file path", "文件路径", "差分包地址", "差分包文件地址"],
    "升级文件的MD5值": ["升级文件的md5值", "升级文件md5值", "升级文件MD5值", "md5", "md5值", "md5 checksum"],
    "文件大小": ["升级包大小", "升级包大小(byte)", "升级包大小（byte）", "升级包大小(Byte)", "升级包大小（Byte）", "文件大小", "升级文件包大小（byte）", "升级文件包大小(byte)", "file size", "size"],
    "语言信息": ["语言信息", "language", "lang", "语言"],
    "品牌": ["品牌", "品牌组", "brand"],
    "区域": ["区域", "区域组", "地区", "地区组", "region", "area", "country", "market", "territory"],
    "SHA256": ["sha256", "sha256值"],
    "EULA文件地址": ["eula文件地址", "eula", "eula url"],
    "MAC": ["mac", "mac地址", "mac address", "场测mac", "场测MAC", "测试mac", "测试MAC", "mac范围", "mac range"],
}

# Rule ID feature: three columns inserted right after "OTA类型" in the generated
# "升级规则" sheet, in this exact order. The order must never change.
RULE_ID_INSERT_AFTER_HEADER = "OTA类型"
RULE_ID_COLUMNS = ["升级方式", "升级日期", "Rule ID"]

# "升级方式" is the human-readable OTA type text; OTA类型 itself holds the numeric
# code (1 / 2 / 11). Map the code back to its text for the 升级方式 column.
OTA_TYPE_CODE_TO_TEXT = {
    "1": "正常升级",
    "2": "强制升级",
    "11": "FTE强制升级",
}

FORWARD_MERGED_CELL_FILL_FIELDS = {"升级包地址", "升级文件的MD5值", "文件大小"}
NEAREST_MERGED_CELL_FILL_FIELDS = {"MAC", "品牌", "区域", "OTA类型"}
MERGED_CELL_FILL_FIELDS = FORWARD_MERGED_CELL_FILL_FIELDS | NEAREST_MERGED_CELL_FILL_FIELDS

FORM_FIELD_ALIASES = {
    "device_code": ["设备类型代码", "设备代码", "device type code", "device code", "devicetypecode"],
    "device_name": ["设备类型名称", "设备名称", "device type name", "device name", "devicetypename"],
    "feature_code": ["feature code", "featurecode", "特征码"],
    "mac_group": ["mac地址组", "mac组", "mac group", "mac address group"],
}

STANDARD_MESSAGE_LANGUAGES = [
    "默认", "简体中文", "英语", "泰语", "意大利语", "荷兰语", "葡萄牙语", "马来语",
    "韩语", "日语", "波兰语", "繁体中文", "德语", "西班牙语", "印尼语",
    "法语", "阿拉伯语", "希伯来语", "芬兰语", "挪威语", "俄语", "瑞典语", "加泰罗尼亚语",
]

MESSAGE_LANGUAGE_NAMES = STANDARD_MESSAGE_LANGUAGES + [
    "简体中文", "中文", "英文", "English", "Thai", "Italian", "Dutch",
    "Portuguese", "Malay", "Korean", "Japanese", "Polish", "German",
    "Spanish", "French", "Arabic", "Hebrew", "Finnish", "Norwegian",
    "Russian", "Swedish", "Catalan", "Indonesian", "Bahasa Indonesia", "Traditional Chinese",
    "中文 (原文)", "中文(原文)", "繁體中文", "日本語", "Deutsch",
    "Italiano", "Nederlands", "Português", "Portugues", "Français", "Francais", "Español", "Espanol",
    "한국어", "Polski", "ไทย", "Bahasa Melayu", "西语", "Spaish", "US-eng", "Mexico-spa",
    "葡语", "葡語", "SA-por", "por", "简中", "簡中", "繁中",
]

MESSAGE_LANGUAGE_ALIASES = {
    "简体中文": "简体中文",
    "中文": "简体中文",
    "中文 (原文)": "简体中文",
    "中文(原文)": "简体中文",
    "简中": "简体中文",
    "簡中": "简体中文",
    "Chinese": "简体中文",
    "Simplified Chinese": "简体中文",
    "英文": "英语",
    "英语": "英语",
    "English": "英语",
    "英语 (English)": "英语",
    "US-eng": "英语",
    "Thai": "泰语",
    "ไทย": "泰语",
    "Italian": "意大利语",
    "Italiano": "意大利语",
    "Dutch": "荷兰语",
    "Nederlands": "荷兰语",
    "Portuguese": "葡萄牙语",
    "Português": "葡萄牙语",
    "Portugues": "葡萄牙语",
    "葡语": "葡萄牙语",
    "葡語": "葡萄牙语",
    "SA-por": "葡萄牙语",
    "por": "葡萄牙语",
    "Malay": "马来语",
    "Bahasa Melayu": "马来语",
    "Korean": "韩语",
    "한국어": "韩语",
    "Japanese": "日语",
    "日本語": "日语",
    "Polish": "波兰语",
    "Polski": "波兰语",
    "German": "德语",
    "Deutsch": "德语",
    "French": "法语",
    "Français": "法语",
    "Francais": "法语",
    "Arabic": "阿拉伯语",
    "Hebrew": "希伯来语",
    "Finnish": "芬兰语",
    "Norwegian": "挪威语",
    "Russian": "俄语",
    "Swedish": "瑞典语",
    "Catalan": "加泰罗尼亚语",
    "Spanish": "西班牙语",
    "Spaish": "西班牙语",
    "西语": "西班牙语",
    "Español": "西班牙语",
    "Espanol": "西班牙语",
    "Mexico-spa": "西班牙语",
    "Indonesian": "印尼语",
    "Bahasa Indonesia": "印尼语",
    "Traditional Chinese": "繁体中文",
    "繁體中文": "繁体中文",
    "繁中": "繁体中文",
}


class OTATemplateApp:

    def __init__(self, root):
        self.root = root
        self.root.title("OTA部署文件生成工具")
        self.root.geometry("1180x960")

        self.config = self.load_config()

        self.source_file_var = tk.StringVar(value=self.config.get("source_file", ""))
        self.device_code_var = tk.StringVar(value=self.config.get("device_code", ""))
        self.device_name_var = tk.StringVar(value=self.config.get("device_name", ""))
        self.feature_code_var = tk.StringVar(value=self.config.get("feature_code", ""))
        self.aws_suffix_var = tk.StringVar(value=self.config.get("aws_suffix", ""))
        self.aws_full_location_var = tk.StringVar(value="")
        self.mac_group_var = tk.StringVar(value=self.config.get("mac_group", ""))
        self.device_history = self.build_initial_device_history(self.config.get("device_history", []))
        self.remember_current_device_history(update_widget=False)

        self.device_code_var.trace_add("write", self.update_aws_suffix_from_device_code)
        self.update_aws_suffix_from_device_code()

        self.status_var = tk.StringVar(value="准备就绪")
        self.download_progress_var = tk.DoubleVar(value=0)
        self.download_status_var = tk.StringVar(value="下载进度：0%")

        self.animation_running = False
        self.animation_dots = 0
        self.animation_text = "正在处理"

        self.rule_id_range_var = tk.StringVar(value="")

        self.log_queue = queue.Queue()
        self.source_wb = None
        self.data_sheet_name = None
        self.paste_text = None
        self.normal_entry_bg = "white"
        self.paste_parse_after_id = None
        self.last_auto_parse_signature = ""

        self.build_ui()
        self.root.after(100, self.process_log_queue)

    # ==========================================================
    # Config
    # ==========================================================

    def load_config(self):
        try:
            if os.path.exists(CONFIG_FILE):
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def get_resource_path(self, filename):
        base_dir = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
        return os.path.join(base_dir, filename)

    def get_user_ota_message_templates_file(self):
        return getattr(self, "user_ota_message_templates_file", USER_OTA_MESSAGE_TEMPLATES_FILE)

    def load_raw_ota_message_templates_file(self, template_path):
        try:
            with open(template_path, "r", encoding="utf-8") as f:
                raw_templates = json.load(f)
        except Exception:
            return {}
        return raw_templates if isinstance(raw_templates, dict) else {}

    def normalize_ota_message_template_options(self, raw_value, source_name="builtin"):
        if isinstance(raw_value, dict) and isinstance(raw_value.get("templates"), list):
            raw_options = raw_value.get("templates", [])
        elif isinstance(raw_value, list):
            if raw_value and all(isinstance(item, dict) and "messages" in item for item in raw_value):
                raw_options = raw_value
            else:
                raw_options = [{"name": "内置模板", "messages": raw_value}]
        else:
            raw_options = []

        options = []
        for index, raw_option in enumerate(raw_options, start=1):
            if isinstance(raw_option, dict) and "messages" in raw_option:
                messages = raw_option.get("messages", [])
                name = str(raw_option.get("name") or f"模板 {index}").strip()
                created_at = str(raw_option.get("created_at") or "").strip()
            else:
                messages = raw_option
                name = f"模板 {index}"
                created_at = ""
            standardized_messages = self.standardize_ota_messages(messages)
            if standardized_messages:
                options.append({
                    "name": name,
                    "created_at": created_at,
                    "source": source_name,
                    "messages": standardized_messages,
                })
        return options

    def load_ota_message_templates(self):
        if hasattr(self, "_ota_message_templates"):
            return self._ota_message_templates

        templates = {}
        for source_name, template_path in (
            ("user", self.get_user_ota_message_templates_file()),
            ("builtin", self.get_resource_path(OTA_MESSAGE_TEMPLATES_FILE)),
        ):
            raw_templates = self.load_raw_ota_message_templates_file(template_path)
            for device_code, raw_value in raw_templates.items():
                code = str(device_code or "").strip()
                if not code:
                    continue
                options = self.normalize_ota_message_template_options(raw_value, source_name=source_name)
                if options:
                    templates.setdefault(code, []).extend(options)

        self._ota_message_templates = templates
        return templates

    def get_ota_message_template_options(self, device_code):
        code = str(device_code or "").strip()
        if not code:
            return []
        templates = self.load_ota_message_templates()
        if code in templates:
            return [
                dict(option, messages=[dict(item) for item in option.get("messages", [])])
                for option in templates[code]
            ]
        normalized_code = self.normalize_header(code)
        for template_code, options in templates.items():
            if self.normalize_header(template_code) == normalized_code:
                return [
                    dict(option, messages=[dict(item) for item in option.get("messages", [])])
                    for option in options
                ]
        return []

    def get_ota_message_template(self, device_code):
        options = self.get_ota_message_template_options(device_code)
        if options:
            return [dict(item) for item in options[0].get("messages", [])]
        return []

    def format_ota_message_template_storage(self, options):
        return {
            "templates": [
                {
                    "name": option.get("name", "模板"),
                    "created_at": option.get("created_at", ""),
                    "messages": option.get("messages", []),
                }
                for option in options
                if option.get("messages")
            ]
        }

    def get_default_ota_message_template_name(self, device_code):
        code = str(device_code or "OTA").strip() or "OTA"
        return f"{code}_{datetime.now().strftime('%Y%m%d_%H%M')}"

    def get_unique_ota_message_template_name(self, requested_name, existing_options):
        name = str(requested_name or "").strip()
        if not name:
            name = "语言提示模板"
        existing_names = {str(option.get("name") or "").strip() for option in existing_options}
        if name not in existing_names:
            return name
        return f"{name} ({datetime.now().strftime('%Y%m%d_%H%M%S')})"

    def prompt_ota_message_template_name(self, device_code, source_messages):
        default_name = self.get_default_ota_message_template_name(device_code)
        if not hasattr(self, "root"):
            return default_name
        language_names = "、".join(item.get("language", "") for item in source_messages[:8] if item.get("language"))
        prompt = (
            f"请为这套语言提示模板命名。\n\n"
            f"设备类型代码：{device_code}\n"
            f"语言数量：{len(source_messages)}\n"
            f"语言：{language_names or '未命名语言'}\n\n"
            "以后选择模板时会显示这个名称。"
        )
        name = simpledialog.askstring(
            "保存语言提示模板",
            prompt,
            initialvalue=default_name,
            parent=self.root,
        )
        return str(name or default_name).strip() or default_name

    def save_ota_message_template_candidate(self, device_code, source_messages, name=None):
        code = str(device_code or "").strip()
        messages = self.standardize_ota_messages(source_messages)
        if not code or not messages:
            return False

        user_template_path = self.get_user_ota_message_templates_file()
        raw_templates = self.load_raw_ota_message_templates_file(user_template_path)
        existing_options = self.normalize_ota_message_template_options(raw_templates.get(code, []), source_name="user")
        message_key = json.dumps(messages, ensure_ascii=False, sort_keys=True)
        template_name = self.get_unique_ota_message_template_name(
            name or self.get_default_ota_message_template_name(code),
            existing_options,
        )
        existing_options = [
            option for option in existing_options
            if json.dumps(option.get("messages", []), ensure_ascii=False, sort_keys=True) != message_key
        ]
        new_option = {
            "name": template_name,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "source": "user",
            "messages": messages,
        }
        raw_templates[code] = self.format_ota_message_template_storage([new_option] + existing_options[:9])
        try:
            with open(user_template_path, "w", encoding="utf-8") as f:
                json.dump(raw_templates, f, ensure_ascii=False, indent=2)
            if hasattr(self, "_ota_message_templates"):
                delattr(self, "_ota_message_templates")
            return True
        except Exception as e:
            self.log_message(f"保存语言提示模板失败：{e}")
            return False

    def select_ota_messages(self, source_messages, device_code):
        source_messages = self.standardize_ota_messages(source_messages)
        if source_messages:
            return source_messages, "source"
        template_messages = self.get_ota_message_template(device_code)
        if template_messages:
            return template_messages, "device_template"
        return self.get_default_ota_messages(), "default"

    def choose_existing_ota_message_template_option(self, device_code, template_options):
        if len(template_options) <= 1:
            return template_options[0] if template_options else None

        dialog = tk.Toplevel(self.root)
        dialog.title("选择语言提示模板")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("560x360")
        dialog.configure(bg="#F6F8FB")

        selected = {"option": None}
        tk.Label(
            dialog,
            text=f"设备类型代码：{device_code}\n请选择要使用的已有语言提示模板",
            bg="#F6F8FB",
            fg="#1F2937",
            font=("Arial", 12, "bold"),
            justify="left",
        ).pack(anchor="w", padx=18, pady=(16, 8))

        listbox = tk.Listbox(dialog, height=8, exportselection=False)
        listbox.pack(fill="x", padx=18)
        for index, option in enumerate(template_options, start=1):
            title = option.get("name") or f"模板 {index}"
            created_at = option.get("created_at") or "无时间"
            language_count = len(option.get("messages", []))
            listbox.insert(tk.END, f"{index}. {title} | {created_at} | {language_count}种语言")
        listbox.selection_set(0)

        preview = tk.Text(dialog, height=7, wrap="word")
        preview.pack(fill="both", expand=True, padx=18, pady=10)

        def refresh_preview(_event=None):
            selection = listbox.curselection()
            if not selection:
                return
            option = template_options[selection[0]]
            lines = []
            for message in option.get("messages", [])[:6]:
                description = str(message.get("description", "")).strip().replace("\n", " ")
                lines.append(f"{message.get('language')}: {description[:90]}")
            preview.config(state="normal")
            preview.delete("1.0", tk.END)
            preview.insert("1.0", "\n".join(lines))
            preview.config(state="disabled")

        def confirm():
            selection = listbox.curselection()
            if selection:
                selected["option"] = template_options[selection[0]]
            dialog.destroy()

        def cancel():
            selected["option"] = None
            dialog.destroy()

        listbox.bind("<<ListboxSelect>>", refresh_preview)
        refresh_preview()

        button_frame = tk.Frame(dialog, bg="#F6F8FB")
        button_frame.pack(fill="x", padx=18, pady=(0, 16))
        tk.Button(button_frame, text="使用选中模板", command=confirm, width=16).pack(side="right", padx=(8, 0))
        tk.Button(button_frame, text="取消", command=cancel, width=10).pack(side="right")
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self.root.wait_window(dialog)
        return selected["option"]

    def choose_ota_message_source(self, device_code, source_messages, template_options):
        result = messagebox.askyesnocancel(
            "选择升级提示语来源",
            (
                f"设备类型代码 {device_code} 已有语言提示模板。\n\n"
                f"本次粘贴内容识别到：{len(source_messages)} 种语言\n"
                f"已有参考模板：{len(template_options)} 套\n\n"
                "选择“是”：使用本次粘贴内容，并保存为新的参考模板。\n"
                "选择“否”：从已有模板中选择。\n"
                "选择“取消”：停止生成。"
            ),
            parent=self.root,
        )
        if result is None:
            return None, None
        if result:
            return "source", None
        selected_option = self.choose_existing_ota_message_template_option(device_code, template_options)
        if not selected_option:
            return None, None
        return "template", selected_option

    def select_ota_messages_for_generation(self, source_messages, device_code, source_type="", prompt_user=True):
        source_messages = self.standardize_ota_messages(source_messages)
        template_options = self.get_ota_message_template_options(device_code)
        if source_messages:
            if source_type == "pasted_text" and device_code:
                selected_source = "source"
                selected_template = None
                if template_options and prompt_user and hasattr(self, "root"):
                    selected_source, selected_template = self.choose_ota_message_source(device_code, source_messages, template_options)
                    if selected_source is None:
                        raise ValueError("用户取消选择升级提示语来源。")
                if selected_source == "template" and selected_template:
                    return [dict(item) for item in selected_template.get("messages", [])], "device_template_selected"
                template_name = None
                if prompt_user and hasattr(self, "root"):
                    template_name = self.prompt_ota_message_template_name(device_code, source_messages)
                saved = self.save_ota_message_template_candidate(device_code, source_messages, name=template_name)
                return source_messages, "source_saved" if saved else "source"
            return source_messages, "source"
        if template_options:
            return [dict(item) for item in template_options[0].get("messages", [])], "device_template"
        return self.get_default_ota_messages(), "default"

    def load_device_history_seed(self):
        if hasattr(self, "_device_history_seed"):
            return self._device_history_seed
        seed_path = self.get_resource_path(DEVICE_HISTORY_SEED_FILE)
        try:
            with open(seed_path, "r", encoding="utf-8") as f:
                raw_seed = json.load(f)
        except Exception:
            raw_seed = []
        self._device_history_seed = self.normalize_device_history(raw_seed)
        return self._device_history_seed

    def build_initial_device_history(self, config_history):
        return self.normalize_device_history(
            list(config_history or []) + list(self.load_device_history_seed())
        )

    def get_default_aws_suffix(self, device_code=None):
        if device_code is None:
            device_code = self.device_code_var.get().strip()
        today = datetime.now().strftime("%Y%m%d")
        return f"/ota/{device_code.strip()}/{today}"

    def get_default_aws_full_location(self, device_code=None):
        """Build the full OTA download location shown in the UI.
        This is informational only and is not editable by the user.
        """
        suffix = self.get_default_aws_suffix(device_code)
        return self.sanitize_url(AWS_PREFIX.rstrip("/") + suffix)

    def update_aws_suffix_from_device_code(self, *args):
        suffix = self.get_default_aws_suffix()
        self.aws_suffix_var.set(suffix)
        self.aws_full_location_var.set(self.sanitize_url(AWS_PREFIX.rstrip("/") + suffix))

    def normalize_device_history(self, history):
        normalized = []
        seen_codes = set()
        for item in history or []:
            if not isinstance(item, dict):
                continue
            device_code = str(item.get("device_code") or item.get("code") or "").strip()
            if not device_code or device_code in seen_codes:
                continue
            seen_codes.add(device_code)
            normalized.append({
                "device_code": device_code,
                "device_name": str(item.get("device_name", "") or "").strip(),
                "feature_code": str(item.get("feature_code", "") or "").strip(),
                "mac_group": str(item.get("mac_group", "") or "").strip(),
            })
            if len(normalized) >= 50:
                break
        return normalized

    def get_device_history_codes(self):
        return [item["device_code"] for item in self.device_history if item.get("device_code")]

    def refresh_device_code_history_widget(self):
        if hasattr(self, "device_code_entry"):
            self.device_code_entry["values"] = self.get_device_history_codes()

    def remember_current_device_history(self, update_widget=True):
        device_code = self.device_code_var.get().strip()
        if not device_code:
            return
        record = {
            "device_code": device_code,
            "device_name": self.device_name_var.get().strip(),
            "feature_code": self.feature_code_var.get().strip(),
            "mac_group": self.mac_group_var.get().strip(),
        }
        remaining = [
            item for item in self.device_history
            if item.get("device_code") != device_code
        ]
        self.device_history = self.normalize_device_history([record] + remaining)
        if update_widget:
            self.refresh_device_code_history_widget()

    def get_device_history_record(self, device_code):
        device_code = str(device_code or "").strip()
        if not device_code:
            return None
        for item in self.device_history:
            if item.get("device_code") == device_code:
                return item
        return None

    def apply_device_history_for_code(self, device_code=None):
        record = self.get_device_history_record(device_code or self.device_code_var.get())
        if not record:
            return False
        self.device_code_var.set(record.get("device_code", ""))
        self.device_name_var.set(record.get("device_name", ""))
        self.feature_code_var.set(record.get("feature_code", ""))
        self.mac_group_var.set(record.get("mac_group", ""))
        self.update_aws_suffix_from_device_code()
        return True

    def on_device_code_history_selected(self, event=None):
        if self.apply_device_history_for_code():
            self.save_config()

    def save_config(self):
        self.update_aws_suffix_from_device_code()
        self.remember_current_device_history()
        source_file = self.source_file_var.get().strip()
        if source_file and os.path.isdir(source_file):
            last_source_dir = source_file
        else:
            last_source_dir = os.path.dirname(source_file) if source_file else ""
        data = {
            "source_file": source_file,
            "last_source_dir": last_source_dir,
            "device_code": self.device_code_var.get().strip(),
            "device_name": self.device_name_var.get().strip(),
            "feature_code": self.feature_code_var.get().strip(),
            "aws_suffix": self.aws_suffix_var.get().strip(),
            "mac_group": self.mac_group_var.get().strip(),
            "device_history": self.device_history,
        }
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self.log_message(f"配置保存失败：{e}")

    # ==========================================================
    # UI
    # ==========================================================

    def build_ui(self):
        self.root.configure(bg="#F6F8FB")
        frame = tk.Frame(self.root, padx=16, pady=14, bg="#F6F8FB")
        frame.pack(fill=tk.BOTH, expand=True)

        label_style = {"bg": "#F6F8FB", "fg": "#27364A", "font": ("Arial", 11, "bold")}
        entry_style = {"bg": "#FFFFFF", "fg": "#1F2937", "relief": tk.SOLID, "bd": 1, "highlightthickness": 1, "highlightbackground": "#D7DEE8"}
        button_style = {"fg": "black", "font": ("Arial", 13, "bold"), "height": 2, "relief": tk.FLAT, "bd": 0, "cursor": "hand2"}

        tk.Label(frame, text="工作目录 / Source Excel：", **label_style).grid(row=0, column=0, sticky="w")
        tk.Entry(frame, textvariable=self.source_file_var, width=90, **entry_style).grid(row=0, column=1, padx=8, sticky="we")
        tk.Button(frame, text="选择", command=self.select_source_path, bg="#E7EDF6", activebackground="#D2DDEC", font=("Arial", 10, "bold"), relief=tk.FLAT, cursor="hand2").grid(row=0, column=2, padx=4, ipadx=12, ipady=3)

        tk.Label(frame, text="邮件/表格内容：", **label_style).grid(row=1, column=0, sticky="nw", pady=8)
        paste_frame = tk.Frame(frame, bg="#F6F8FB")
        paste_frame.grid(row=1, column=1, columnspan=2, sticky="nsew", padx=8, pady=6)
        self.paste_text = scrolledtext.ScrolledText(
            paste_frame,
            height=9,
            width=105,
            wrap="word",
            bg="#FFFFFF",
            fg="#111827",
            insertbackground="#111827",
            relief=tk.SOLID,
            bd=1,
            font=("Menlo", 10),
        )
        self.paste_text.grid(row=0, column=0, sticky="nsew")
        self.paste_text.tag_configure("field_label_found", background="#D9EAD3", foreground="#1F5E2C")
        self.paste_text.tag_configure("field_value_found", background="#B6D7A8", foreground="#123D1E")
        self.paste_text.tag_configure("field_label_missing", background="#FFF2CC", foreground="#7A5C00")
        self.paste_text.bind("<<Paste>>", self.schedule_pasted_text_parse)
        self.paste_text.bind("<KeyRelease>", self.schedule_pasted_text_parse)
        paste_button_frame = tk.Frame(paste_frame, bg="#F6F8FB")
        paste_button_frame.grid(row=0, column=1, sticky="ne", padx=(10, 0))
        tk.Button(
            paste_button_frame,
            text="清除内容",
            command=self.clear_pasted_text,
            bg="#FCE5CD",
            fg="black",
            font=("Arial", 10, "bold"),
            relief=tk.FLAT,
            cursor="hand2"
        ).pack(fill=tk.X, ipadx=10, ipady=4)
        paste_frame.columnconfigure(0, weight=1)
        paste_frame.rowconfigure(0, weight=1)

        tk.Label(frame, text="设备类型代码：", **label_style).grid(row=2, column=0, sticky="w", pady=6)
        self.device_code_entry = ttk.Combobox(
            frame,
            textvariable=self.device_code_var,
            width=43,
            values=self.get_device_history_codes(),
        )
        self.device_code_entry.bind("<<ComboboxSelected>>", self.on_device_code_history_selected)
        self.device_code_entry.bind("<Return>", self.on_device_code_history_selected)
        self.device_code_entry.bind("<FocusOut>", self.on_device_code_history_selected)
        self.device_code_entry.grid(row=2, column=1, sticky="w", padx=6)

        tk.Label(frame, text="设备类型名称：", **label_style).grid(row=3, column=0, sticky="w", pady=6)
        self.device_name_entry = tk.Entry(frame, textvariable=self.device_name_var, width=45, **entry_style)
        self.device_name_entry.grid(row=3, column=1, sticky="w", padx=6)

        tk.Label(frame, text="Feature Code / 特征码：", **label_style).grid(row=4, column=0, sticky="w", pady=6)
        self.feature_code_entry = tk.Entry(frame, textvariable=self.feature_code_var, width=45, **entry_style)
        self.feature_code_entry.grid(row=4, column=1, sticky="w", padx=6)

        tk.Label(frame, text="MAC 地址组：", **label_style).grid(row=5, column=0, sticky="w", pady=6)
        self.mac_group_entry = tk.Entry(frame, textvariable=self.mac_group_var, width=70, **entry_style)
        self.mac_group_entry.grid(row=5, column=1, sticky="w", padx=6)
        self.normal_entry_bg = self.device_code_entry.cget("background")

        tk.Label(frame, text="AWS 文件路径：", **label_style).grid(row=6, column=0, sticky="w", pady=6)

        aws_display_label = tk.Label(
            frame,
            textvariable=self.aws_full_location_var,
            fg="#0B57D0",
            bg="#EAF2FF",
            relief=tk.SOLID,
            bd=1,
            anchor="w",
            padx=8,
            pady=4,
            width=82
        )
        aws_display_label.grid(row=6, column=1, sticky="we", padx=6)

        rule_id_frame = tk.Frame(frame, bg="#F6F8FB")
        rule_id_frame.grid(row=7, column=0, columnspan=3, sticky="we", pady=(4, 0))
        tk.Label(rule_id_frame, text="Rule ID 范围：", **label_style).pack(side="left")
        tk.Entry(
            rule_id_frame, textvariable=self.rule_id_range_var, width=24, **entry_style
        ).pack(side="left", padx=(6, 6))
        tk.Label(
            rule_id_frame,
            text="单个如 65213，或范围如 65210-65213",
            bg="#F6F8FB", fg="#7A8699", font=("Arial", 9),
        ).pack(side="left")
        tk.Button(
            rule_id_frame,
            text="更新Rule ID列",
            command=self.update_rule_id_columns,
            bg="#C9DAF8", activebackground="#A4C2F4",
            fg="black", font=("Arial", 10, "bold"),
            relief=tk.FLAT, cursor="hand2",
        ).pack(side="left", padx=(12, 0), ipadx=10, ipady=3)

        button_frame = tk.Frame(frame, bg="#F6F8FB")
        button_frame.grid(row=8, column=0, columnspan=3, pady=18)
        inner_button_frame = tk.Frame(button_frame, bg="#F6F8FB")
        inner_button_frame.pack()

        self.generate_template_button = tk.Button(
            inner_button_frame, text="生成模版文件", command=self.generate_template_only,
            bg="#F9CB63", activebackground="#F4B183", width=20, **button_style
        )
        self.generate_template_button.grid(row=0, column=0, padx=12)

        self.download_zip_button = tk.Button(
            inner_button_frame, text="下载/上传ZIP升级包", command=self.download_upload_zip_package,
            bg="#9FC5E8", activebackground="#6FA8DC", width=22, **button_style
        )
        self.download_zip_button.grid(row=0, column=1, padx=12)

        self.auto_execute_button = tk.Button(
            inner_button_frame, text="自动执行", command=self.auto_execute_all,
            bg="#93C47D", activebackground="#6AA84F", width=20, **button_style
        )
        self.auto_execute_button.grid(row=0, column=2, padx=12)

        tk.Label(button_frame, textvariable=self.status_var, bg="#F6F8FB", fg="#415064", font=("Arial", 11, "bold")).pack(pady=8)

        progress_frame = tk.Frame(frame, bg="#F6F8FB")
        progress_frame.grid(row=9, column=0, columnspan=3, sticky="we", pady=8)
        tk.Label(progress_frame, textvariable=self.download_status_var, bg="#F6F8FB", font=("Arial", 11), fg="#333333").pack()
        self.download_progress_bar = ttk.Progressbar(
            progress_frame, variable=self.download_progress_var, maximum=100, length=700, mode="determinate"
        )
        self.download_progress_bar.pack(pady=4)

        tk.Label(frame, text="执行日志：", **label_style).grid(row=10, column=0, sticky="nw", pady=6)
        self.log_text = scrolledtext.ScrolledText(frame, height=24, width=145, bg="#0F172A", fg="#E5E7EB", insertbackground="#E5E7EB", relief=tk.FLAT, font=("Menlo", 10))
        self.log_text.grid(row=10, column=1, columnspan=2, sticky="nsew", padx=6)

        copy_log_button = tk.Button(
            frame,
            text="复制日志 / Copy Log",
            command=self.copy_log_to_clipboard,
            bg="#D9EAD3",
            fg="black",
            font=("Arial", 10, "bold"),
            relief=tk.FLAT,
            cursor="hand2"
        )
        copy_log_button.grid(row=11, column=1, sticky="e", padx=6, pady=6)

        self.log_text.bind("<Command-a>", self.select_all_log_text)
        self.log_text.bind("<Control-a>", self.select_all_log_text)

        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(10, weight=1)

    # ==========================================================
    # Center Popup
    # ==========================================================

    def show_center_message(self, title, message):
        popup = tk.Toplevel(self.root)
        popup.title(title)
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)

        width = 460
        height = 200
        self.root.update_idletasks()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        x = root_x + (root_width // 2) - (width // 2)
        y = root_y + (root_height // 2) - (height // 2)
        popup.geometry(f"{width}x{height}+{x}+{y}")

        frame = tk.Frame(popup, padx=22, pady=22)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text=title, font=("Arial", 16, "bold"), fg="#2F5597").pack(pady=(0, 14))
        tk.Label(frame, text=message, font=("Arial", 12), justify="center", wraplength=400).pack(pady=6)
        tk.Button(
            frame, text="确定", width=12, bg="#93C47D", fg="black",
            font=("Arial", 11, "bold"), command=popup.destroy, cursor="hand2"
        ).pack(pady=16)
        popup.focus_force()

    def show_generated_file_prompt(self, output_path):
        open_file = messagebox.askyesno(
            "文件已生成",
            f"文件已生成：\n{output_path}\n\n是否打开 Excel 文件？",
            parent=self.root,
        )
        if open_file:
            self.open_excel_file(output_path)

    def open_excel_file(self, output_path):
        try:
            if os.name == "posix":
                subprocess.Popen(["open", output_path])
            elif os.name == "nt":
                os.startfile(output_path)
            else:
                subprocess.Popen(["xdg-open", output_path])
            self.log_message(f"已调用系统打开 Excel 文件：{output_path}")
        except Exception as e:
            self.log_message(f"打开 Excel 文件失败：{e}")
            self.show_center_error("打开失败", f"文件已生成，但打开失败：\n{e}")

    def show_center_error(self, title, message):
        popup = tk.Toplevel(self.root)
        popup.title(title)
        popup.transient(self.root)
        popup.grab_set()
        popup.resizable(False, False)

        width = 520
        height = 260
        self.root.update_idletasks()
        root_x = self.root.winfo_x()
        root_y = self.root.winfo_y()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()
        x = root_x + (root_width // 2) - (width // 2)
        y = root_y + (root_height // 2) - (height // 2)
        popup.geometry(f"{width}x{height}+{x}+{y}")

        frame = tk.Frame(popup, padx=20, pady=20)
        frame.pack(fill="both", expand=True)
        tk.Label(frame, text=title, font=("Arial", 16, "bold"), fg="red").pack(pady=(0, 12))
        text_box = tk.Text(frame, height=7, wrap="word", font=("Consolas", 10))
        text_box.pack(fill="both", expand=True)
        text_box.insert("1.0", message)
        text_box.config(state="disabled")
        tk.Button(
            frame, text="关闭", width=12, bg="#E06666", fg="black",
            font=("Arial", 11, "bold"), command=popup.destroy, cursor="hand2"
        ).pack(pady=12)
        popup.focus_force()

    # ==========================================================
    # Button / animation / log
    # ==========================================================

    def disable_all_buttons(self):
        self.generate_template_button.config(state=tk.DISABLED)
        self.download_zip_button.config(state=tk.DISABLED)
        self.auto_execute_button.config(state=tk.DISABLED)

    def enable_all_buttons(self):
        self.generate_template_button.config(state=tk.NORMAL)
        self.download_zip_button.config(state=tk.NORMAL)
        self.auto_execute_button.config(state=tk.NORMAL)

    def start_animation(self, text="正在处理"):
        self.animation_running = True
        self.animation_dots = 0
        self.animation_text = text
        self.animate_status()

    def stop_animation(self, final_text="完成"):
        self.animation_running = False
        self.status_var.set(final_text)
        self.enable_all_buttons()

    def animate_status(self):
        if not self.animation_running:
            return
        dots = "." * (self.animation_dots % 4)
        self.status_var.set(f"{self.animation_text}{dots}")
        self.animation_dots += 1
        self.root.after(400, self.animate_status)

    def log_message(self, message):
        """Write a detailed timestamped message into the UI execution log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        text = str(message)
        # Keep section dividers visually clean.
        if not text or set(text.strip()) <= {"="}:
            line = text
        else:
            line = f"[{timestamp}] {text}"
        self.log_text.insert(tk.END, line + "\n")
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def copy_log_to_clipboard(self):
        """Copy all content in the execution log window to the system clipboard."""
        content = self.log_text.get("1.0", tk.END).strip()
        if not content:
            self.show_center_message("提示", "执行日志为空，没有可复制内容。")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(content)
        self.root.update()
        self.show_center_message("复制成功", "执行日志已复制到剪贴板。")

    def select_all_log_text(self, event=None):
        self.log_text.tag_add("sel", "1.0", tk.END)
        self.log_text.mark_set("insert", "1.0")
        self.log_text.see("insert")
        return "break"

    def thread_log(self, message):
        self.log_queue.put(("log", message))

    def thread_progress(self, percent, text=None):
        self.log_queue.put(("progress", percent, text))

    def process_log_queue(self):
        try:
            while True:
                item = self.log_queue.get_nowait()
                if item[0] == "log":
                    self.log_message(item[1])
                elif item[0] == "progress":
                    percent = max(0, min(100, float(item[1])))
                    text = item[2]
                    self.download_progress_var.set(percent)
                    self.download_status_var.set(text if text else f"下载进度：{percent:.1f}%")
                elif item[0] == "done":
                    success_count = item[1]
                    self.stop_animation("ZIP 下载完成")
                    self.download_progress_var.set(100)
                    self.download_status_var.set("全部 ZIP 下载任务完成")
                    self.show_center_message("下载完成", f"ZIP 文件下载完成。\n\n成功下载：{success_count} 个文件")
                elif item[0] == "s3_done":
                    success_count = item[1]
                    fail_count = item[2]
                    if fail_count == 0:
                        self.stop_animation("S3 上传完成")
                        self.show_center_message("上传完成", f"所有文件已上传至 S3。\n\n成功：{success_count} 个文件")
                    else:
                        self.stop_animation("S3 上传部分失败")
                        self.show_center_error("上传结果", f"S3 上传完成（有部分失败）。\n\n成功：{success_count} 个\n失败：{fail_count} 个\n\n请查看执行日志了解详情。")
                elif item[0] == "error":
                    error_message = item[1]
                    self.stop_animation("任务失败")
                    self.show_center_error("错误", error_message)
        except queue.Empty:
            pass
        self.root.after(100, self.process_log_queue)

    # ==========================================================
    # File select
    # ==========================================================

    def select_source_path(self):
        choose_file = messagebox.askyesnocancel(
            "选择路径",
            "请选择要设置的路径类型：\n\n是：选择 Source Excel 文件\n否：选择工作目录\n取消：不选择",
            parent=self.root,
        )
        if choose_file is None:
            return
        if choose_file:
            self.select_source_file()
        else:
            self.select_working_directory()

    def select_source_file(self):
        initial_dir = self.config.get("last_source_dir", "")
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = os.path.expanduser("~")
        file_path = filedialog.askopenfilename(
            title="选择 Source Excel 文件",
            initialdir=initial_dir,
            filetypes=[("Excel Files", "*.xlsx *.xlsm"), ("All Files", "*.*")]
        )
        if file_path:
            self.source_file_var.set(file_path)
            self.config["last_source_dir"] = os.path.dirname(file_path)
            self.save_config()
            self.log_message(f"已选择 Source Excel 文件：{file_path}")

    def select_working_directory(self):
        initial_dir = self.config.get("last_source_dir", "")
        current_path = self.source_file_var.get().strip()
        if current_path and os.path.isdir(current_path):
            initial_dir = current_path
        elif current_path and os.path.isfile(current_path):
            initial_dir = os.path.dirname(current_path)
        if not initial_dir or not os.path.exists(initial_dir):
            initial_dir = os.path.expanduser("~")
        folder_path = filedialog.askdirectory(
            title="选择工作目录",
            initialdir=initial_dir,
        )
        if folder_path:
            self.source_file_var.set(folder_path)
            self.config["last_source_dir"] = folder_path
            self.save_config()
            self.log_message(f"已选择工作目录：{folder_path}")

    def get_pasted_source_text(self):
        if not self.paste_text:
            return ""
        return self.paste_text.get("1.0", tk.END).strip()

    def clear_pasted_text(self):
        self.paste_text.delete("1.0", tk.END)
        self.clear_auto_fill_highlights()
        self.clear_paste_text_highlights()
        self.last_auto_parse_signature = ""
        self.log_message("已清空邮件/表格粘贴内容。")

    def schedule_pasted_text_parse(self, event=None):
        if self.paste_parse_after_id:
            try:
                self.root.after_cancel(self.paste_parse_after_id)
            except Exception:
                pass
        self.paste_parse_after_id = self.root.after(300, self.auto_parse_pasted_text)

    def auto_parse_pasted_text(self):
        self.paste_parse_after_id = None
        text = self.get_pasted_source_text()
        if not text:
            self.clear_auto_fill_highlights()
            self.clear_paste_text_highlights()
            self.last_auto_parse_signature = ""
            return
        signature = str(hash(text))
        if signature == self.last_auto_parse_signature:
            return
        self.last_auto_parse_signature = signature
        self.parse_pasted_text_to_form(show_popup=False, show_log=True)

    def clear_paste_text_highlights(self):
        if not self.paste_text:
            return
        for tag_name in ("field_label_found", "field_value_found", "field_label_missing"):
            self.paste_text.tag_remove(tag_name, "1.0", tk.END)

    def highlight_text_occurrences(self, needle, tag_name):
        needle = str(needle or "").strip()
        if not needle or not self.paste_text:
            return
        start = "1.0"
        while True:
            pos = self.paste_text.search(needle, start, tk.END, nocase=True)
            if not pos:
                break
            end = f"{pos}+{len(needle)}c"
            self.paste_text.tag_add(tag_name, pos, end)
            start = end

    def highlight_pasted_content(self, values):
        self.clear_paste_text_highlights()
        field_alias_groups = {
            "device_code": FORM_FIELD_ALIASES["device_code"],
            "device_name": FORM_FIELD_ALIASES["device_name"],
            "feature_code": FORM_FIELD_ALIASES["feature_code"],
            "mac_group": FORM_FIELD_ALIASES["mac_group"],
        }
        for field, aliases in field_alias_groups.items():
            tag_name = "field_label_found" if values.get(field) else "field_label_missing"
            for alias in aliases:
                self.highlight_text_occurrences(alias, tag_name)
            if values.get(field):
                self.highlight_text_occurrences(values[field], "field_value_found")

    def split_pasted_line(self, line):
        raw_line = str(line or "").strip("\r\n")
        stripped_line = raw_line.strip()
        if "\t" in raw_line:
            return [cell.strip() for cell in raw_line.split("\t")]
        if "|" in stripped_line:
            return [cell.strip() for cell in stripped_line.split("|") if cell.strip()]
        return [cell.strip() for cell in re.split(r"\s{2,}", stripped_line) if cell.strip()]

    def find_labeled_value(self, text, aliases):
        for alias in aliases:
            escaped = re.escape(alias)
            patterns = [
                rf"(?im)^\s*{escaped}\s*[:：=]\s*(.+?)\s*$",
                rf"(?im)^.*?{escaped}[^:：=\n]{{0,30}}[:：=]\s*(.+?)\s*$",
                rf"(?im)^\s*{escaped}\s*\t+\s*(.+?)\s*$",
                rf"(?im)^\s*{escaped}\s+(.+?)\s*$",
            ]
            for pattern in patterns:
                match = re.search(pattern, text)
                if match:
                    value = match.group(1).strip()
                    value = re.split(r"\s{3,}|\t", value)[0].strip()
                    if value:
                        return value
        return ""

    def extract_form_fields_from_text(self, text):
        values = {}
        for field, aliases in FORM_FIELD_ALIASES.items():
            value = self.find_labeled_value(text, aliases)
            if value:
                values[field] = value
        table_values = self.extract_form_fields_from_table_rows(text)
        values.update({key: value for key, value in table_values.items() if value})
        return values

    def extract_form_fields_from_table_rows(self, text):
        values = {}
        lines = [line for line in str(text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n") if line.strip()]
        table_rows = [self.split_pasted_line(line) for line in lines]
        # The header-table path is the reliable one for horizontal tables (it
        # reads the header row, then pulls each field from the aligned DATA row).
        # Fields it resolves must NOT be overwritten by the label/value fallback
        # below, which only makes sense for VERTICAL "label<TAB>value" pairs.
        header_table_values = self.extract_form_fields_from_header_table_rows(table_rows)
        values.update(header_table_values)
        for cells in table_rows:
            if len(cells) < 2:
                continue
            # Skip a horizontal TABLE HEADER row: if the row's own cells are
            # themselves known header names (e.g. 设备类型代码 / featurecode / 机型),
            # then cells[1] is the next column's header, not this label's value.
            # Treating it as a label/value pair mis-assigns header text as data
            # (bug: 设备类型代码 header -> value 'featurecode').
            if self.is_table_header_row(cells):
                continue
            label = cells[0].strip()
            value = next((cell.strip() for cell in cells[1:] if cell.strip()), "")
            if not label or not value:
                continue
            label_norm = self.normalize_header(label)
            if "featurecode" in label_norm or "特征码" in label_norm:
                if "feature_code" not in header_table_values:
                    values["feature_code"] = value
            elif "设备类型代码" in label_norm or "设备代码" in label_norm or "devicetypecode" in label_norm or "devicecode" in label_norm:
                if "device_code" not in header_table_values:
                    values["device_code"] = value
            elif "设备类型名称" in label_norm or "设备名称" in label_norm or "devicetypename" in label_norm or "devicename" in label_norm:
                if "device_name" not in header_table_values:
                    values["device_name"] = value
            elif label_norm in ("mac组", "mac地址组", "macgroup", "macaddressgroup"):
                if "mac_group" not in header_table_values:
                    values["mac_group"] = value
        return values

    def is_table_header_row(self, cells):
        """True if a pasted row is a horizontal table HEADER row (several cells
        are known column/form headers), as opposed to a vertical label/value
        pair. Used to stop the label/value fallback from mis-reading a header
        row's adjacent header cell as a field value."""
        known = 0
        for cell in cells:
            text = str(cell or "").strip()
            if not text:
                continue
            if self.get_form_field_from_header(text):
                known += 1
                continue
            norm = self.normalize_header(text)
            for aliases in COLUMN_ALIASES.values():
                if any(self.normalize_header(a) == norm for a in aliases):
                    known += 1
                    break
        return known >= 2

    def get_form_field_from_header(self, header):
        header_norm = self.normalize_header(header)
        if "featurecode" in header_norm or "特征码" in header_norm:
            return "feature_code"
        if (
            "设备类型代码" in header_norm
            or "设备代码" in header_norm
            or "devicetypecode" in header_norm
            or "devicecode" in header_norm
        ):
            return "device_code"
        if (
            "设备类型名称" in header_norm
            or "设备名称" in header_norm
            or "devicetypename" in header_norm
            or "devicename" in header_norm
        ):
            return "device_name"
        if header_norm in ("mac组", "mac地址组", "macgroup", "macaddressgroup"):
            return "mac_group"
        return ""

    def is_valid_form_field_value(self, field, value):
        value = str(value or "").strip()
        if not value:
            return False
        if field == "feature_code":
            return self.looks_like_feature_code(value)
        if field == "mac_group":
            return self.looks_like_mac_group(value) or bool(value)
        return True

    def extract_form_fields_from_header_table_rows(self, table_rows):
        values = {}
        for row_index, header_cells in enumerate(table_rows):
            header_mapping = {}
            for col_idx, header in enumerate(header_cells):
                field = self.get_form_field_from_header(header)
                if field:
                    header_mapping[field] = col_idx
            if not header_mapping:
                continue

            # Case A — VERTICAL metadata pair: the row is a "标签<TAB>值" line
            # (e.g. ['设备类型代码（方案号）', 'MT9900NAAnd4KUHD']), NOT a horizontal
            # table header. Here the value sits in the SAME row right after the
            # header cell. Resolve from this row and do NOT scan later rows, so a
            # language line like ['美国-英语/US-eng', ...] can't be mistaken for
            # the device code.
            if not self.is_table_header_row(header_cells):
                for field, col_idx in header_mapping.items():
                    if field in values:
                        continue
                    value = next(
                        (c.strip() for c in header_cells[col_idx + 1:] if c.strip()), ""
                    )
                    if value and self.is_valid_form_field_value(field, value):
                        values[field] = value
                if values:
                    return values
                continue

            # Case B — HORIZONTAL table header: the value for each header column
            # lives in a SUBSEQUENT data row at the same column index.
            for data_cells in table_rows[row_index + 1:]:
                if self.is_metadata_row(data_cells) or self.is_message_table_header_row(data_cells):
                    continue
                if len(data_cells) < 2:
                    continue
                for field, col_idx in header_mapping.items():
                    if field in values:
                        continue
                    if col_idx < len(data_cells):
                        value = data_cells[col_idx].strip()
                        if self.is_valid_form_field_value(field, value):
                            values[field] = value
                if values:
                    return values
        return values

    def clear_auto_fill_highlights(self):
        for entry in self.get_form_entries().values():
            try:
                entry.config(
                    bg=self.normal_entry_bg,
                    highlightthickness=1,
                    highlightbackground="#D9D9D9",
                    highlightcolor="#D9D9D9",
                )
            except tk.TclError:
                pass

    def get_form_entries(self):
        return {
            "device_code": self.device_code_entry,
            "device_name": self.device_name_entry,
            "feature_code": self.feature_code_entry,
            "mac_group": self.mac_group_entry,
        }

    def highlight_auto_filled_fields(self, values):
        extracted_bg = "#D9EAD3"
        missing_bg = "#FFF2CC"
        extracted_border = "#6AA84F"
        missing_border = "#D6A100"
        for field, entry in self.get_form_entries().items():
            try:
                if values.get(field):
                    entry.config(bg=extracted_bg, highlightthickness=2, highlightbackground=extracted_border, highlightcolor=extracted_border)
                else:
                    entry.config(bg=missing_bg, highlightthickness=2, highlightbackground=missing_border, highlightcolor=missing_border)
            except tk.TclError:
                pass

    def apply_pasted_form_fields(self, text, show_log=True):
        values = self.extract_form_fields_from_text(text)
        updated = []
        if values.get("device_code"):
            self.device_code_var.set(values["device_code"])
            updated.append(f"设备类型代码={values['device_code']}")
        if values.get("device_name"):
            self.device_name_var.set(values["device_name"])
            updated.append(f"设备类型名称={values['device_name']}")
        if values.get("feature_code"):
            self.feature_code_var.set(self.validate_feature_code(values["feature_code"]))
            updated.append(f"Feature Code={self.feature_code_var.get()}")
        if values.get("mac_group"):
            self.mac_group_var.set(values["mac_group"])
            updated.append(f"MAC组={values['mac_group']}")
        if values.get("device_code"):
            self.update_aws_suffix_from_device_code()
            self.root.update_idletasks()
        self.highlight_auto_filled_fields(values)
        if show_log:
            if updated:
                self.log_message("自动回填：" + "，".join(updated))
            else:
                self.log_message("未在粘贴内容中识别到可自动回填的设备类型代码、设备名称、Feature Code 或 MAC组。")
        return values, updated

    def normalize_message_language(self, language):
        language = str(language or "").strip()
        language_norm = self.normalize_header(language)
        if not language_norm:
            return ""
        # Exact (normalized) equality: a cell is only recognized as a language when
        # it IS a language label, never when a long description merely contains a
        # language substring (which caused false positives like "por" inside
        # "importante").
        for standard in STANDARD_MESSAGE_LANGUAGES:
            if self.normalize_header(standard) == language_norm:
                return standard
        for alias, standard in MESSAGE_LANGUAGE_ALIASES.items():
            if self.normalize_header(alias) == language_norm:
                return standard
        # Bounded "wrapped label" fallback: handle short labels that embed the
        # language name, e.g. "English (英语)", "英语 (English)", "语种/Language".
        # Only applies to short cells so a long sentence containing a token cannot
        # match.
        if len(language_norm) <= 24:
            tokens = re.split(r"[()/／|\s（）\-－—–]+", language)
            for token in tokens:
                token = token.strip()
                if not token:
                    continue
                token_norm = self.normalize_header(token)
                if not token_norm:
                    continue
                for standard in STANDARD_MESSAGE_LANGUAGES:
                    if self.normalize_header(standard) == token_norm:
                        return standard
                for alias, standard in MESSAGE_LANGUAGE_ALIASES.items():
                    if self.normalize_header(alias) == token_norm:
                        return standard
        return ""

    def is_message_language_name(self, text):
        return bool(self.normalize_message_language(text))

    def extract_language_messages_from_table_rows(self, table_rows):
        messages = []
        for row_index, cells in enumerate(table_rows):
            normalized = [self.normalize_header(cell) for cell in cells]
            language_col = None
            description_col = None
            for col_idx, header in enumerate(normalized):
                if header in ("语种", "语言", "language", "lang") or "语种" in header or "language" in header or ("语言" in header and "language" in header):
                    language_col = col_idx
                if (
                    header in ("描述", "内容", "提示语", "升级提示语", "message", "description", "desc", "content")
                    or ("内容" in header and "content" in header)
                    or "reminder" in header
                ):
                    description_col = col_idx
            if description_col is None:
                continue
            # BUG 1: some source sheets have a description header ("升级提示语") but no
            # language-column header (e.g. first row is [空, 升级提示语] and col0 holds the
            # language names). Fall back to treating column 0 as the language column, but
            # only when the data rows below really look like a language sheet, so this does
            # not mis-fire on OTA data tables whose col0 is a 机型 code.
            if language_col is None:
                if description_col == 0:
                    continue
                language_hits = 0
                for probe_cells in table_rows[row_index + 1:]:
                    if not probe_cells:
                        continue
                    probe = str(probe_cells[0]).strip() if len(probe_cells) else ""
                    if probe and self.normalize_message_language(probe):
                        language_hits += 1
                if language_hits < 2:
                    continue
                language_col = 0
            if language_col == description_col:
                continue

            current_language = ""
            current_description_lines = []

            def flush_current():
                nonlocal current_language, current_description_lines
                if current_language and any(line.strip() for line in current_description_lines):
                    description = "\n".join(line for line in current_description_lines if line.strip()).strip()
                    if description:
                        messages.append({"language": current_language, "description": description})
                current_language = ""
                current_description_lines = []

            for data_cells in table_rows[row_index + 1:]:
                if not any(str(cell).strip() for cell in data_cells):
                    continue

                # BUG 2: a new language entry may ONLY start when the LANGUAGE COLUMN of a
                # physical row holds a valid, standalone language name. Every other non-empty
                # row (bullet continuation lines copied from a multi-line cell, left-shifted
                # fields, etc.) is a continuation of the current language and is appended to
                # the current description. A language is never inferred from body text.
                language_cell = data_cells[language_col].strip() if language_col < len(data_cells) else ""
                language = self.normalize_message_language(language_cell)
                if language:
                    flush_current()
                    current_language = language
                    description = data_cells[description_col].strip() if description_col < len(data_cells) else ""
                    if description:
                        current_description_lines.append(description)
                    continue

                if current_language:
                    if self.is_message_table_header_row(data_cells):
                        flush_current()
                        continue
                    if self.is_metadata_row(data_cells):
                        flush_current()
                        continue
                    if len(data_cells) >= 3:
                        _header_index, _mapping, score = self.detect_columns_from_text_rows([data_cells])
                        if score >= 2:
                            flush_current()
                            continue
                    continuation = "\t".join(cell.strip() for cell in data_cells if str(cell).strip()).strip()
                    if continuation:
                        current_description_lines.append(continuation)
            flush_current()
            break
        return self.dedupe_language_messages(messages)

    def is_message_table_header_row(self, cells):
        normalized = [self.normalize_header(cell) for cell in cells]
        language_cols = [
            idx for idx, header in enumerate(normalized)
            if header in ("语种", "语言", "language", "lang") or "语种" in header or "language" in header or ("语言" in header and "language" in header)
        ]
        content_cols = [
            idx for idx, header in enumerate(normalized)
            if (
                header in ("描述", "内容", "提示语", "升级提示语", "message", "description", "desc", "content")
                or ("内容" in header and "content" in header)
                or "reminder" in header
            )
        ]
        has_language = any(
            header in ("语种", "语言", "language", "lang") or "语种" in header or "language" in header or ("语言" in header and "language" in header)
            for header in normalized
        )
        has_content = any(
            header in ("描述", "内容", "提示语", "升级提示语", "message", "description", "desc", "content")
            or ("内容" in header and "content" in header)
            or "reminder" in header
            for header in normalized
        )
        has_index = any(header in ("序号", "no", "no.", "index") for header in normalized)
        has_distinct_columns = any(language_col != content_col for language_col in language_cols for content_col in content_cols)
        return has_language and has_content and has_distinct_columns and (has_index or len([h for h in normalized if h]) <= 4)

    def extract_language_messages_from_blocks(self, text):
        lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        language_pattern = "|".join(re.escape(name) for name in sorted(MESSAGE_LANGUAGE_NAMES, key=len, reverse=True))
        label_regex = re.compile(rf"^\s*({language_pattern})\s*[:：=]\s*(.*)\s*$", re.IGNORECASE)
        messages = []
        current_language = None
        current_lines = []

        def flush_current():
            nonlocal current_language, current_lines
            if current_language and any(line.strip() for line in current_lines):
                description = "\n".join(current_lines).strip()
                if description:
                    messages.append({
                        "language": self.normalize_message_language(current_language),
                        "description": description,
                    })
            current_language = None
            current_lines = []

        for line in lines:
            match = label_regex.match(line)
            if match:
                flush_current()
                current_language = match.group(1).strip()
                first_content = match.group(2).strip()
                current_lines = [first_content] if first_content else []
                continue
            if current_language:
                cells = self.split_pasted_line(line)
                if len(cells) >= 3:
                    _header_index, _mapping, score = self.detect_columns_from_text_rows([cells])
                    if score >= 2:
                        flush_current()
                        continue
                # Stop if a new non-message data section starts.
                line_norm = self.normalize_header(line)
                if line_norm in ("设备类型代码", "devicecode", "devicetypecode", "featurecode") or line.strip().startswith("http"):
                    flush_current()
                    continue
                current_lines.append(line.rstrip())
        flush_current()
        return self.dedupe_language_messages(messages)

    def dedupe_language_messages(self, messages):
        deduped = []
        seen_languages = set()
        for item in messages:
            language = self.normalize_message_language(item.get("language", ""))
            description = str(item.get("description", "")).strip()
            if not language or not description:
                continue
            # Keep a single row per language. Different source regions that map to
            # the same standard language (e.g. 墨西哥-西语 and 中南美-西语 both ->
            # 西班牙语) collapse to the first occurrence.
            if language in seen_languages:
                continue
            seen_languages.add(language)
            deduped.append({"language": language, "description": description})
        return self.order_ota_messages(deduped)

    def extract_language_messages_from_text(self, text, table_rows=None):
        table_messages = self.extract_language_messages_from_table_rows(table_rows or [])
        block_messages = self.extract_language_messages_from_blocks(text)
        return self.dedupe_language_messages(table_messages + block_messages)

    def parse_pasted_text_to_form(self, show_popup=True, show_log=True):
        text = self.get_pasted_source_text()
        if not text:
            if show_popup:
                self.show_center_message("提示", "请先粘贴邮件或表格内容。")
            return

        _values, updated = self.apply_pasted_form_fields(text, show_log=False)
        self.highlight_pasted_content(_values)

        try:
            source_json = self.build_source_json_from_text(text, use_gui_log=False)
            row_count = len(source_json.get("rows", []))
            message_count = len(source_json.get("messages", []))
            mapping_text = self.format_mapping_for_log(source_json.get("mapping", {}))
        except Exception as e:
            row_count = 0
            message_count = 0
            mapping_text = "未识别到 OTA 数据字段"
            if show_log:
                self.log_message(f"粘贴内容暂未形成 OTA 数据行：{e}")
        if show_log:
            self.log_message("已自动解析粘贴内容。")
            if updated:
                self.log_message("自动回填：" + "，".join(updated))
            else:
                self.log_message("未在粘贴内容中识别到可自动回填的设备类型代码、设备名称、Feature Code 或 MAC组。")
            self.log_message(f"从粘贴内容中识别到 OTA 数据行数：{row_count}")
            self.log_message(f"从粘贴内容中识别到升级提示语数量：{message_count}")
            self.log_message(f"粘贴内容字段映射：{mapping_text}")
        self.save_config()
        if show_popup:
            self.show_center_message("解析完成", f"已解析粘贴内容。\n\n识别到 OTA 数据行数：{row_count}\n识别到升级提示语：{message_count} 条")

    # ==========================================================
    # Column detection
    # ==========================================================

    def normalize_header(self, value):
        if value is None:
            return ""
        text = str(value).strip().lower()
        text = text.replace(" ", "").replace("_", "").replace("-", "")
        text = text.replace("（", "(").replace("）", ")")
        return text

    def detect_columns_by_header(self, ws):
        header_row = None
        best_score = -1
        best_mapping = {}
        for row_idx in range(1, min(ws.max_row, 20) + 1):
            raw_headers = [ws.cell(row=row_idx, column=col).value for col in range(1, ws.max_column + 1)]
            normalized_headers = [self.normalize_header(h) for h in raw_headers]
            mapping = {}
            for target_field, aliases in COLUMN_ALIASES.items():
                for alias in aliases:
                    alias_norm = self.normalize_header(alias)
                    for col_idx, header_norm in enumerate(normalized_headers, start=1):
                        if not header_norm:
                            continue
                        if header_norm == alias_norm or alias_norm in header_norm or header_norm in alias_norm:
                            mapping[target_field] = col_idx
                            break
                    if target_field in mapping:
                        break
            if len(mapping) > best_score:
                best_score = len(mapping)
                header_row = row_idx
                best_mapping = mapping
        return header_row, best_mapping, best_score

    def looks_like_version(self, value):
        text = str(value or "").strip()
        return bool(
            re.match(r"^\d{4}[A-Z]{1,2}\d{4}$", text, re.IGNORECASE)
            or re.match(r"^V\d+(?:\.\d+)+$", text, re.IGNORECASE)
        )

    def looks_like_version_change(self, value):
        text = str(value or "").strip()
        if not re.search(r"_to_", text, re.IGNORECASE):
            return False
        parts = re.split(r"_to_", text, flags=re.IGNORECASE)
        if len(parts) != 2:
            return False
        return all(
            self.looks_like_version(part) or re.match(r"^[A-Z]{1,2}\d{4}$", part, re.IGNORECASE)
            for part in parts
        )

    def looks_like_md5(self, value):
        text = str(value or "").strip()
        return bool(re.match(r"^[a-fA-F0-9]{32}$", text))

    def looks_like_size(self, value):
        text = str(value or "").strip()
        clean_text = re.sub(r"[,\s_]", "", text)
        return bool(re.match(r"^\d{6,15}$", clean_text))

    def looks_like_url_zip(self, value):
        text = str(value or "").strip().lower()
        return "://" in text and ".zip" in text

    def looks_like_model(self, value):
        text = str(value or "").strip()
        if re.search(r"\d{5,}$", text):
            return False
        return bool(re.match(r"^\d{2}[A-Z0-9]+(?:\(.+\))?$", text))

    def looks_like_feature_code(self, value):
        text = str(value or "").strip()
        return bool(re.match(r"^\d{18,}$", text))

    def looks_like_extension(self, value):
        text = str(value or "").strip()
        if (
            not text
            or self.looks_like_model(text)
            or self.looks_like_version(text)
            or self.looks_like_version_change(text)
            or self.looks_like_md5(text)
            or self.looks_like_size(text)
            or self.looks_like_feature_code(text)
            or self.extract_zip_filename(text)
        ):
            return False
        return bool(re.match(r"^[A-Z0-9][A-Z0-9()_-]{3,}$", text, re.IGNORECASE))

    def looks_like_mac_group(self, value):
        text = str(value or "").strip()
        if not text:
            return False
        compact_hex_groups = re.findall(r"(?i)\b[a-f0-9]{8}\b", text)
        return bool(
            re.fullmatch(r"(?i)[a-f0-9]{8,12}", text)
            or
            re.search(r"(?i)[a-f0-9]{8}", text)
            and (":" in text or "-" in text or "到" in text or "、" in text or len(compact_hex_groups) >= 2)
        )

    def looks_like_ota_type_text(self, value):
        text = str(value or "").strip()
        normalized = text.replace(" ", "").replace("－", "-").replace("–", "-").replace("—", "-")
        return bool(
            re.fullmatch(r"(1|2|11)(?:\.0+)?", normalized)
            or normalized in ("正常", "强制", "非强制")
            or "正常升级" in normalized
            or "强制升级" in normalized
            or "非强制升级" in normalized
            or "FTE强制升级" in normalized
        )

    def normalize_brand_value(self, value):
        text = str(value or "").strip()
        if not text:
            return ""
        known_brands = ["Hisense", "Toshiba", "VIDAA", "Noblex", "Polytron"]
        for brand in known_brands:
            if re.search(rf"(?i)\b{re.escape(brand)}\b", text):
                return brand
        return ""

    def looks_like_region_value(self, value):
        text = str(value or "").strip()
        if not text:
            return False
        normalized = text.upper()
        return normalized in ["US", "USA", "EU", "UK", "CA", "AU", "APAC", "LATAM", "MENA", "GLOBAL"] or text in ["全球", "不限制", "不限", "不区分", "欧洲", "北美"]

    def detect_columns_without_header(self, ws):
        scores = {}
        for col_idx in range(1, ws.max_column + 1):
            col_scores = {
                "机型信息": 0, "机器扩展码": 0, "OTA类型": 0, "源版本": 0, "目标版本": 0,
                "版本变化": 0, "升级包地址": 0, "升级文件的MD5值": 0, "文件大小": 0, "品牌": 0, "区域": 0, "MAC": 0,
            }
            for row_idx in range(1, min(ws.max_row, 30) + 1):
                text = str(ws.cell(row=row_idx, column=col_idx).value or "").strip()
                if not text:
                    continue
                if self.looks_like_model(text):
                    col_scores["机型信息"] += 3
                if self.looks_like_feature_code(text):
                    col_scores["机器扩展码"] += 1
                if self.looks_like_extension(text):
                    col_scores["机器扩展码"] += 3
                if text in ["正常升级", "强制升级", "非强制升级", "FTE强制升级"]:
                    col_scores["OTA类型"] += 5
                if self.looks_like_version(text):
                    col_scores["源版本"] += 2
                    col_scores["目标版本"] += 2
                if "_to_" in text:
                    col_scores["版本变化"] += 5
                if self.looks_like_url_zip(text):
                    col_scores["升级包地址"] += 6
                if self.looks_like_md5(text):
                    col_scores["升级文件的MD5值"] += 6
                if self.looks_like_size(text):
                    col_scores["文件大小"] += 4
                if text.lower() in ["hisense", "toshiba", "vidaa"]:
                    col_scores["品牌"] += 5
                if text.upper() in ["US", "USA", "EU", "UK", "CA", "AU", "APAC", "LATAM", "MENA", "GLOBAL"] or text.lower() in ["north america", "europe", "global"]:
                    col_scores["区域"] += 4
                if ("、" in text or "," in text) and re.search(r"[a-fA-F0-9]{8}", text):
                    col_scores["MAC"] += 3
            scores[col_idx] = col_scores

        mapping = {}
        used_cols = set()
        field_priority = ["升级包地址", "升级文件的MD5值", "文件大小", "OTA类型", "版本变化", "品牌", "区域", "MAC", "机型信息", "机器扩展码"]
        for field in field_priority:
            best_col = None
            best_score = 0
            for col_idx, col_scores in scores.items():
                if col_idx in used_cols:
                    continue
                if col_scores.get(field, 0) > best_score:
                    best_score = col_scores[field]
                    best_col = col_idx
            if best_col and best_score > 0:
                mapping[field] = best_col
                used_cols.add(best_col)

        version_cols = []
        for col_idx, col_scores in scores.items():
            if col_idx in used_cols:
                continue
            version_score = max(col_scores.get("源版本", 0), col_scores.get("目标版本", 0))
            if version_score > 0:
                version_cols.append(col_idx)
        version_cols = sorted(version_cols)
        if len(version_cols) >= 1:
            mapping["源版本"] = version_cols[0]
        if len(version_cols) >= 2:
            mapping["目标版本"] = version_cols[1]
        return 0, mapping, len(mapping)

    # ==========================================================
    # Excel read
    # ==========================================================

    def read_source_excel(self, file_path, use_gui_log=True):
        if use_gui_log:
            self.log_message("正在读取 Excel...")
            self.log_message(f"Source Excel 路径：{file_path}")
            self.log_message(f"文件大小：{os.path.getsize(file_path)} bytes")
        self.source_wb = load_workbook(file_path, data_only=False)
        if use_gui_log:
            self.log_message("Excel 工作簿打开成功。")
            self.log_message(f"工作表数量：{len(self.source_wb.worksheets)}")
            self.log_message("工作表列表：" + ", ".join([ws.title for ws in self.source_wb.worksheets]))

        best_ws = None
        best_header_row = None
        best_mapping = {}
        best_score = -1
        best_has_header = True

        for ws in self.source_wb.worksheets:
            header_row, mapping, score = self.detect_columns_by_header(ws)
            no_header_row, no_header_mapping, no_header_score = self.detect_columns_without_header(ws)
            if use_gui_log:
                self.log_message(f"检测工作表：{ws.title}")
                self.log_message(f"  尺寸：{ws.max_row} 行 x {ws.max_column} 列")
                self.log_message(f"  表头识别字段数量：{score}")
                self.log_message(f"  表头模式字段映射：{self.format_mapping_for_log(mapping)}")
                self.log_message(f"  无表头识别字段数量：{no_header_score}")
                self.log_message(f"  无表头模式字段映射：{self.format_mapping_for_log(no_header_mapping)}")
            if no_header_score > score:
                current_score = no_header_score
                current_mapping = no_header_mapping
                current_header_row = no_header_row
                current_has_header = False
            else:
                current_score = score
                current_mapping = mapping
                current_header_row = header_row
                current_has_header = True
            if current_score > best_score:
                best_ws = ws
                best_header_row = current_header_row
                best_mapping = current_mapping
                best_score = current_score
                best_has_header = current_has_header

        if best_ws is None or best_score <= 0:
            raise ValueError("未识别到有效工作表。")

        self.data_sheet_name = best_ws.title
        start_row = best_header_row + 1 if best_has_header else 1

        if use_gui_log:
            self.log_message(f"最终选择数据工作表：{best_ws.title}")
            self.log_message(f"识别模式：{'有表头模式' if best_has_header else '无表头自动识别模式'}")
            if best_has_header:
                self.log_message(f"表头所在行：{best_header_row}")
            else:
                self.log_message("源文件未识别到表头：数据将从第 1 行开始读取。")
            self.log_message("字段映射结果：")
            for field, col_idx in best_mapping.items():
                self.log_message(f"  {field} -> {get_column_letter(col_idx)}列 / 示例值：{best_ws.cell(start_row, col_idx).value}")
            self.log_required_field_status(best_mapping)

        data_rows = []
        for row_idx in range(start_row, best_ws.max_row + 1):
            row_data = {}
            has_any_value = False

            # Keep a copy of the whole original Excel row.
            # This is important because some customer source files use changing
            # header names for the FTP ZIP URL column, for example 整包地址 / 整包文件 / 升级包地址.
            # Even if the normal field mapping misses that column, generation can still
            # scan this row and extract the exact .zip filename.
            all_cell_values = []
            for col_idx in range(1, best_ws.max_column + 1):
                raw_value = self.get_excel_cell_value(best_ws, row_idx, col_idx)
                if raw_value not in (None, ""):
                    has_any_value = True
                    all_cell_values.append(str(raw_value).strip())

            for field, col_idx in best_mapping.items():
                value = self.get_excel_cell_value(best_ws, row_idx, col_idx)
                row_data[field] = "" if value is None else str(value).strip()

            if has_any_value:
                if self.is_metadata_row(all_cell_values):
                    continue
                row_data["_row_number"] = row_idx
                row_data["_all_cell_values"] = all_cell_values
                data_rows.append(row_data)

        if use_gui_log:
            self.log_message(f"读取到的数据行数：{len(data_rows)}")
            self.log_message("Excel 读取完成。")
        return data_rows, best_mapping

    def get_excel_cell_value(self, ws, row_idx, col_idx):
        value = ws.cell(row=row_idx, column=col_idx).value
        if value not in (None, ""):
            return value
        for merged_range in ws.merged_cells.ranges:
            if (
                merged_range.min_row <= row_idx <= merged_range.max_row
                and merged_range.min_col <= col_idx <= merged_range.max_col
            ):
                return ws.cell(row=merged_range.min_row, column=merged_range.min_col).value
        return value

    def detect_columns_from_text_rows(self, rows):
        best_header_index = None
        best_mapping = {}
        best_score = -1
        for row_index, cells in enumerate(rows[:20]):
            normalized_headers = [self.normalize_header(cell) for cell in cells]
            mapping = {}
            for target_field, aliases in COLUMN_ALIASES.items():
                for alias in aliases:
                    alias_norm = self.normalize_header(alias)
                    for col_idx, header_norm in enumerate(normalized_headers, start=1):
                        if not header_norm:
                            continue
                        if header_norm == alias_norm or alias_norm in header_norm or header_norm in alias_norm:
                            mapping[target_field] = col_idx
                            break
                    if target_field in mapping:
                        break
            if len(mapping) > best_score:
                best_score = len(mapping)
                best_header_index = row_index
                best_mapping = mapping
        return best_header_index, best_mapping, best_score

    def find_message_table_start_index(self, table_rows, start_index=0):
        for row_index, cells in enumerate(table_rows[start_index:], start=start_index):
            if self.is_message_table_header_row(cells):
                return row_index
        return None

    def fill_merged_cells_in_text_rows(self, table_rows, start_index, mapping):
        """Fill blank pasted-table cells caused by Excel merged cells.

        When merged Excel cells are copied from email/Excel, the pasted text often
        contains the merged value only once, while the surrounding rows are blank.
        For shared fields such as URL, MD5, file size, MAC, brand, region, and
        OTA type, use the nearest visible value to fill blank rows before
        converting to JSON.
        """
        if not table_rows or not mapping:
            return table_rows, []

        max_width = max(len(row) for row in table_rows)
        filled_rows = [list(row) + [""] * (max_width - len(row)) for row in table_rows]
        changed_fields = []
        data_row_indices = [
            row_idx
            for row_idx in range(start_index, len(filled_rows))
            if not self.is_metadata_row(filled_rows[row_idx])
            and not self.is_message_table_header_row(filled_rows[row_idx])
        ]
        if not data_row_indices:
            return filled_rows, changed_fields

        for field in MERGED_CELL_FILL_FIELDS:
            col_idx = mapping.get(field)
            if not col_idx:
                continue
            col = col_idx - 1
            non_empty = []
            for row_idx in data_row_indices:
                value = filled_rows[row_idx][col].strip()
                if value:
                    non_empty.append((row_idx, value))
            if not non_empty:
                continue

            unique_values = list(dict.fromkeys(value for _row_idx, value in non_empty))
            should_fill = len(unique_values) == 1 or len(non_empty) <= max(3, len(data_row_indices) // 2)
            if not should_fill:
                continue

            changed_count = 0
            if field in FORWARD_MERGED_CELL_FILL_FIELDS:
                last_value = ""
                first_value = non_empty[0][1]
                for row_idx in data_row_indices:
                    current_value = filled_rows[row_idx][col].strip()
                    if current_value:
                        last_value = current_value
                        continue
                    fill_value = last_value or first_value
                    if fill_value:
                        filled_rows[row_idx][col] = fill_value
                        changed_count += 1
                if changed_count:
                    changed_fields.append(f"{field}({changed_count}行)")
                continue

            for row_idx in data_row_indices:
                if filled_rows[row_idx][col].strip():
                    continue
                nearest_row, nearest_value = min(non_empty, key=lambda item: abs(item[0] - row_idx))
                if nearest_value:
                    filled_rows[row_idx][col] = nearest_value
                    changed_count += 1
            if changed_count:
                changed_fields.append(f"{field}({changed_count}行)")

        return filled_rows, changed_fields

    def classify_pasted_cell_field(self, value):
        text = str(value or "").strip()
        if not text:
            return None, ""
        if self.looks_like_feature_code(text):
            return "特征码", text
        if self.extract_zip_filename(text):
            return "升级包地址", text
        if self.looks_like_md5(text):
            return "升级文件的MD5值", text
        if self.looks_like_size(text):
            return "文件大小", text
        if self.looks_like_version_change(text):
            return "版本变化", text
        if self.looks_like_version(text):
            return "版本", text
        if self.looks_like_mac_group(text):
            return "MAC", text
        brand = self.normalize_brand_value(text)
        if brand:
            return "品牌", brand
        if self.looks_like_ota_type_text(text):
            return "OTA类型", text
        if self.looks_like_region_value(text):
            return "区域", text
        if self.looks_like_extension(text):
            return "机器扩展码", text
        if self.looks_like_model(text):
            return "机型信息", text
        return None, ""

    def pasted_row_matches_header_mapping(self, cells, mapping):
        """Return True when the row already lines up with the detected header.

        Content-based realignment is only meant to repair rows whose columns were
        shifted (merged cells copied without blanks, left-shifted fields, etc.).
        For a clean tab-separated paste the header mapping is already correct, and
        realigning by content corrupts the row (e.g. a model code that does not
        match the Chinese ``32DM66GE`` pattern, or a trailing ``最大检测数`` number
        that looks like a model). Trust the header when the mapped identity and
        version columns are self-consistent.
        """
        if not mapping:
            return False

        def value(field):
            col_idx = mapping.get(field)
            if not col_idx or col_idx - 1 >= len(cells):
                return ""
            return str(cells[col_idx - 1] or "").strip()

        model = value("机型信息")
        extension = value("机器扩展码")
        source_version = value("源版本")
        target_version = value("目标版本")

        has_identity = bool(model) and bool(extension) and model != extension
        has_versions = self.looks_like_version(source_version) and self.looks_like_version(target_version)
        if not (has_identity and has_versions):
            return False

        # Guard against the trailing-number-as-model case: a real model/extension
        # must not be a bare number (those come from 最大检测数 / 文件大小 columns).
        if re.fullmatch(r"\d+", model) or re.fullmatch(r"\d+", extension):
            return False
        return True

    def realign_pasted_row_by_content(self, cells, mapping):
        if not mapping:
            return cells

        # Only repair rows that are actually misaligned. A clean paste whose row
        # already matches the detected header must be kept verbatim.
        if self.pasted_row_matches_header_mapping(cells, mapping):
            return cells

        semantic_fields = {
            "机型信息", "机器扩展码", "源版本", "目标版本", "版本变化",
            "升级包地址", "升级文件的MD5值", "文件大小", "MAC", "品牌", "区域", "OTA类型",
        }
        max_col = max(max(mapping.values(), default=0), len(cells))
        aligned = list(cells) + [""] * (max_col - len(cells))

        # Clear semantic columns first; then place values by their actual content.
        # This fixes rows where copied merged cells removed blank columns and shifted
        # fields into wrong positions.
        for field in semantic_fields:
            col_idx = mapping.get(field)
            if col_idx and col_idx <= len(aligned):
                aligned[col_idx - 1] = ""

        version_values = []
        for value in cells:
            field, normalized_value = self.classify_pasted_cell_field(value)
            if field == "版本":
                version_values.append(normalized_value)
                continue
            if not field or field == "特征码" or field not in mapping:
                continue
            col_idx = mapping[field] - 1
            if col_idx < 0:
                continue
            if col_idx >= len(aligned):
                aligned.extend([""] * (col_idx + 1 - len(aligned)))
            if not aligned[col_idx]:
                aligned[col_idx] = normalized_value

        for field, version_value in zip(("源版本", "目标版本"), version_values):
            col_idx = mapping.get(field)
            if not col_idx:
                continue
            col_idx -= 1
            if col_idx >= len(aligned):
                aligned.extend([""] * (col_idx + 1 - len(aligned)))
            if not aligned[col_idx]:
                aligned[col_idx] = version_value

        return aligned

    def realign_pasted_table_by_content(self, table_rows, start_index, mapping):
        if not table_rows or not mapping:
            return table_rows
        aligned_rows = []
        for row_index, cells in enumerate(table_rows):
            if row_index >= start_index:
                if self.is_metadata_row(cells) or self.is_message_table_header_row(cells):
                    aligned_rows.append(cells)
                else:
                    aligned_rows.append(self.realign_pasted_row_by_content(cells, mapping))
            else:
                aligned_rows.append(cells)
        return aligned_rows

    def merge_wrapped_mac_rows_in_text_table(self, table_rows, start_index, mapping):
        """Merge rows split by a newline inside the copied MAC cell."""
        mac_col_idx = mapping.get("MAC")
        if not table_rows or not mac_col_idx:
            return table_rows, 0

        mac_col = mac_col_idx - 1
        merged_rows = []
        merged_count = 0
        for row_index, cells in enumerate(table_rows):
            current_cells = list(cells)
            if row_index >= start_index and merged_rows and current_cells:
                first_cell = str(current_cells[0] or "").strip()
                previous_row = merged_rows[-1]
                previous_has_ota_data = (
                    not self.is_metadata_row(previous_row)
                    and not self.is_message_table_header_row(previous_row)
                    and any(str(cell).strip() for cell in previous_row)
                )
                current_is_mac_continuation = (
                    self.looks_like_mac_group(first_cell)
                    and not self.looks_like_model(first_cell)
                    and not self.looks_like_version(first_cell)
                    and not self.looks_like_version_change(first_cell)
                    and not self.extract_zip_filename(first_cell)
                )
                if previous_has_ota_data and current_is_mac_continuation:
                    if len(previous_row) <= mac_col:
                        previous_row.extend([""] * (mac_col + 1 - len(previous_row)))
                    existing_mac = str(previous_row[mac_col] or "").strip()
                    previous_row[mac_col] = " ".join(value for value in (existing_mac, first_cell) if value)

                    for offset, value in enumerate(current_cells[1:], start=1):
                        target_col = mac_col + offset
                        if len(previous_row) <= target_col:
                            previous_row.extend([""] * (target_col + 1 - len(previous_row)))
                        if str(value or "").strip() and not str(previous_row[target_col] or "").strip():
                            previous_row[target_col] = str(value).strip()
                    merged_count += 1
                    continue
            merged_rows.append(current_cells)
        return merged_rows, merged_count

    def has_ota_rule_anchor(self, cells):
        if not cells or self.is_metadata_row(cells) or self.is_message_table_header_row(cells):
            return False

        fields = []
        for cell in cells:
            field, _value = self.classify_pasted_cell_field(cell)
            if field:
                fields.append(field)

        has_extension = "机器扩展码" in fields
        has_model = "机型信息" in fields
        version_count = fields.count("版本")
        has_version_change = "版本变化" in fields
        has_package = any(field in fields for field in ("升级包地址", "升级文件的MD5值", "文件大小"))

        if has_extension:
            return True
        if has_model and (version_count > 0 or has_version_change or has_package):
            return True
        if version_count >= 2 or has_version_change:
            return True
        return False

    def extract_zip_urls_from_text(self, text):
        pattern = r"(?i)\b(?:ftp|https?)://[^\s<>'\"，。；;]+?\.zip(?:[^\s<>'\"，。；;]*)?"
        urls = [match.group(0).strip() for match in re.finditer(pattern, text or "")]
        return list(dict.fromkeys(urls))

    def extract_text_field_fallbacks(self, text):
        fields = {}
        for field, aliases in COLUMN_ALIASES.items():
            value = self.find_labeled_value(text, aliases)
            if value:
                fields[field] = value

        zip_urls = self.extract_zip_urls_from_text(text)
        if zip_urls and not fields.get("升级包地址"):
            fields["升级包地址"] = zip_urls[0]

        sha_matches = re.findall(r"(?i)\b[a-f0-9]{64}\b", text or "")
        if sha_matches and not fields.get("SHA256"):
            fields["SHA256"] = sha_matches[0]

        md5_matches = re.findall(r"(?i)\b[a-f0-9]{32}\b", text or "")
        if md5_matches and not fields.get("升级文件的MD5值"):
            fields["升级文件的MD5值"] = md5_matches[0]

        version_matches = re.findall(r"\b\d{4}[A-Za-z]{1,2}\d{4}\b", text or "")
        if version_matches and not fields.get("源版本"):
            fields["源版本"] = version_matches[0]
        if len(version_matches) >= 2 and not fields.get("目标版本"):
            fields["目标版本"] = version_matches[1]

        version_change_match = re.search(r"\b[\w.-]+_to_[\w.-]+\b", text or "", re.IGNORECASE)
        if version_change_match and not fields.get("版本变化"):
            fields["版本变化"] = version_change_match.group(0)

        size_match = re.search(r"(?i)(?:file\s*size|size|文件大小|升级包大小)[^\d]{0,20}([\d,\s_]{6,})", text or "")
        if size_match and not fields.get("文件大小"):
            fields["文件大小"] = size_match.group(1).strip()

        return fields, zip_urls

    def read_source_text(self, text, use_gui_log=True):
        text = str(text or "").strip()
        if not text:
            raise ValueError("请先粘贴邮件或表格内容。")

        if use_gui_log:
            self.log_message("正在解析粘贴的邮件/表格内容...")
            self.log_message(f"粘贴内容长度：{len(text)} 字符")

        lines = [line for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n") if line.strip()]
        table_rows = [self.split_pasted_line(line) for line in lines]
        table_rows = [row for row in table_rows if len(row) >= 2]
        header_index, mapping, score = self.detect_columns_from_text_rows(table_rows)

        data_rows = []
        skipped_metadata_count = 0
        if score > 0 and header_index is not None:
            start_index = header_index + 1
            table_rows, wrapped_mac_merge_count = self.merge_wrapped_mac_rows_in_text_table(table_rows, start_index, mapping)
            message_start_index = self.find_message_table_start_index(table_rows, start_index)
            if message_start_index is not None:
                skipped_metadata_count += len(table_rows[message_start_index:])
                table_rows_for_rules = table_rows[:message_start_index]
            else:
                table_rows_for_rules = table_rows

            rule_anchor_indices = {
                row_idx
                for row_idx in range(start_index, len(table_rows_for_rules))
                if self.has_ota_rule_anchor(table_rows_for_rules[row_idx])
            }

            table_rows_for_rules = self.realign_pasted_table_by_content(table_rows_for_rules, start_index, mapping)
            table_rows_for_rules, merged_fill_fields = self.fill_merged_cells_in_text_rows(table_rows_for_rules, start_index, mapping)
            if use_gui_log:
                self.log_message("粘贴内容识别为表格模式。")
                self.log_message(f"表头所在行：{header_index + 1}")
                self.log_message(f"字段映射结果：{self.format_mapping_for_log(mapping)}")
                self.log_message(f"原始表格中识别到 OTA 规则锚点行数：{len(rule_anchor_indices)}")
                if wrapped_mac_merge_count:
                    self.log_message(f"检测到 MAC 单元格内换行，已合并续行：{wrapped_mac_merge_count} 行")
                if message_start_index is not None:
                    self.log_message(f"检测到升级提示语表头，OTA 规则解析已在第 {message_start_index + 1} 行前停止。")
                if merged_fill_fields:
                    self.log_message("检测到疑似合并单元格空白，已自动补齐字段：" + "，".join(merged_fill_fields))

            in_message_table = False
            for row_index in range(start_index, len(table_rows_for_rules)):
                cells = table_rows_for_rules[row_index]
                row_offset = row_index + 1
                if not any(str(cell).strip() for cell in cells):
                    continue
                if self.is_message_table_header_row(cells):
                    skipped_metadata_count += 1
                    in_message_table = True
                    continue
                if in_message_table:
                    skipped_metadata_count += 1
                    continue
                if self.is_metadata_row(cells):
                    skipped_metadata_count += 1
                    continue
                if row_index not in rule_anchor_indices:
                    skipped_metadata_count += 1
                    continue
                cells = self.realign_pasted_row_by_content(cells, mapping)
                row_data = {}
                for field, col_idx in mapping.items():
                    cell_index = col_idx - 1
                    row_data[field] = cells[cell_index].strip() if cell_index < len(cells) else ""
                if not self.is_ota_source_row(row_data):
                    continue
                row_data["_row_number"] = row_offset
                row_data["_all_cell_values"] = [cell.strip() for cell in cells if str(cell).strip()]
                data_rows.append(row_data)
        else:
            fields, zip_urls = self.extract_text_field_fallbacks(text)
            mapping = {field: index + 1 for index, field in enumerate(fields.keys())}
            source_values = [value for value in fields.values() if str(value).strip()]
            if zip_urls:
                for index, zip_url in enumerate(zip_urls, start=1):
                    row_data = dict(fields)
                    row_data["升级包地址"] = zip_url
                    row_data["_row_number"] = index
                    row_data["_all_cell_values"] = list(dict.fromkeys(source_values + [zip_url]))
                    data_rows.append(row_data)
            elif fields:
                row_data = dict(fields)
                row_data["_row_number"] = 1
                row_data["_all_cell_values"] = source_values
                data_rows.append(row_data)

            if use_gui_log:
                self.log_message("粘贴内容识别为普通文本模式。")
                self.log_message(f"字段映射结果：{', '.join(mapping.keys()) if mapping else '未识别到任何字段'}")

        if not data_rows:
            raise ValueError("未能从粘贴内容中识别到 OTA 数据。请确认内容包含表格字段、版本信息或 ZIP 地址。")

        if use_gui_log:
            self.log_required_field_status(mapping)
            if skipped_metadata_count:
                self.log_message(f"已跳过非 OTA 数据行数量：{skipped_metadata_count}（例如设备类型代码 / featurecode / 语言提示语表，不作为 OTA 规则）")
            self.log_message(f"读取到的数据行数：{len(data_rows)}")
            self.log_message("粘贴内容解析完成。")
        return data_rows, mapping

    def is_ota_source_row(self, row_data):
        values = [str(value or "").strip() for value in row_data.values()]
        if self.is_metadata_row(values):
            return False
        model = str(row_data.get("机型信息", "")).strip()
        extension = str(row_data.get("机器扩展码", "")).strip()
        has_identity = self.looks_like_model(model) or self.looks_like_extension(extension)
        if not has_identity:
            return False

        version_values = [
            row_data.get("源版本", ""),
            row_data.get("目标版本", ""),
        ]
        has_version = any(self.looks_like_version(value) for value in version_values)
        has_version_change = self.looks_like_version_change(row_data.get("版本变化", ""))
        has_package = bool(
            self.extract_zip_filename(row_data.get("升级包地址", ""))
            or self.looks_like_md5(row_data.get("升级文件的MD5值", ""))
            or self.looks_like_size(row_data.get("文件大小", ""))
        )
        return has_version or has_version_change or has_package

    def is_metadata_row(self, values):
        joined = " ".join(values)
        normalized = self.normalize_header(joined)
        metadata_markers = [
            "featurecode",
            "特征码",
            "设备类型代码",
            "设备代码",
            "devicetypecode",
            "devicecode",
            "设备类型名称",
            "devicename",
        ]
        return any(marker in normalized for marker in metadata_markers)

    def extract_language_messages_from_excel(self, file_path):
        wb = load_workbook(file_path, data_only=True, read_only=True)
        messages = []
        try:
            for ws in wb.worksheets:
                rows = []
                for row in ws.iter_rows(values_only=True):
                    cells = ["" if value is None else str(value).strip() for value in row]
                    if any(cells):
                        rows.append(cells)
                sheet_messages = self.extract_language_messages_from_table_rows(rows)
                if sheet_messages:
                    messages.extend(sheet_messages)
        finally:
            wb.close()
        return self.dedupe_language_messages(messages)

    def build_source_json_from_text(self, text, use_gui_log=True):
        rows, mapping = self.read_source_text(text, use_gui_log=use_gui_log)
        lines = [line for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n") if line.strip()]
        table_rows = [self.split_pasted_line(line) for line in lines]
        source_json = {
            "source_type": "pasted_text",
            "source_path": "",
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "form_fields": self.extract_form_fields_from_text(text),
            "mapping": mapping,
            "rows": rows,
            "messages": self.extract_language_messages_from_text(text, table_rows),
        }
        if use_gui_log:
            self.log_source_json_summary(source_json)
        return source_json

    def build_source_json_from_excel(self, file_path, use_gui_log=True):
        rows, mapping = self.read_source_excel(file_path, use_gui_log=use_gui_log)
        source_json = {
            "source_type": "excel",
            "source_path": file_path,
            "created_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "form_fields": {},
            "mapping": mapping,
            "rows": rows,
            "messages": self.extract_language_messages_from_excel(file_path),
        }
        if use_gui_log:
            self.log_source_json_summary(source_json)
        return source_json

    def build_source_json(self, source_file_path="", pasted_text="", use_gui_log=True):
        if pasted_text:
            return self.build_source_json_from_text(pasted_text, use_gui_log=use_gui_log)
        return self.build_source_json_from_excel(source_file_path, use_gui_log=use_gui_log)

    def log_source_json_summary(self, source_json):
        self.log_message("已转换为 JSON 中间格式。")
        self.log_message(f"JSON 数据来源：{source_json.get('source_type')}")
        self.log_message(f"JSON OTA 数据行数：{len(source_json.get('rows', []))}")
        self.log_message(f"JSON 升级提示语数量：{len(source_json.get('messages', []))}")

    def generate_json_output_path(self, excel_output_path):
        base, _ext = os.path.splitext(excel_output_path)
        return base + ".json"

    def save_source_json(self, source_json, output_path):
        json_path = self.generate_json_output_path(output_path)
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(source_json, f, ensure_ascii=False, indent=2)
        return json_path

    # ==========================================================
    # FTP
    # ==========================================================

    def get_unique_ftp_urls(self, rows):
        """Return unique FTP ZIP URLs from parsed source rows.

        This keeps compatibility with the normal column mapping path.
        The source Excel may call this field 整包地址 / 升级包地址 / 升级文件地址.
        """
        ftp_urls = []
        for row in rows:
            url = self.get_value(row, "升级包地址")
            if url and str(url).strip().lower().startswith("ftp://") and ".zip" in str(url).lower():
                ftp_urls.append(str(url).strip())
        return list(dict.fromkeys(ftp_urls))

    def scan_ftp_zip_urls_from_excel(self, source_file_path):
        """Scan the whole workbook for FTP .zip URLs.

        This is intentionally independent from header detection.  Some source
        files use different column names such as 整包地址, and the safest
        behavior for the download step is to scan every cell and extract any
        ftp://...zip value.
        """
        self.thread_log("开始扫描 Source Excel 中所有单元格，查找 FTP ZIP 地址...")
        self.thread_log(f"扫描文件：{source_file_path}")
        wb = load_workbook(source_file_path, data_only=True, read_only=True)
        urls = []
        for ws in wb.worksheets:
            self.thread_log(f"扫描工作表：{ws.title}，尺寸：{ws.max_row} 行 x {ws.max_column} 列")
            for row in ws.iter_rows():
                for cell in row:
                    value = cell.value
                    if value is None:
                        continue
                    text = str(value).strip()
                    if text.lower().startswith("ftp://") and ".zip" in text.lower():
                        urls.append(text)
                        self.thread_log(f"发现 FTP ZIP 地址：工作表={ws.title}, 单元格={cell.coordinate}")
                        self.thread_log(f"  原始地址：{text}")
        wb.close()
        unique_urls = list(dict.fromkeys(urls))
        self.thread_log(f"扫描完成：发现 FTP ZIP 地址 {len(urls)} 条，去重后 {len(unique_urls)} 个文件。")
        return unique_urls

    def parse_ftp_url(self, ftp_url):
        """Parse FTP URL from source Excel.

        Excel usually stores URLs such as:
        ftp://ftpuser@10.18.203.204/upload/.../32S60S%289A%29_User/.../USBOTA_xxx.zip

        Important: urlparse keeps the path percent-encoded.  FTP servers store
        the real directory name with characters such as parentheses, not with
        literal %28/%29.  Therefore we must unquote the path before cwd/RETR.
        Otherwise ftp.cwd('/upload/.../32S60S%289A%29_User/...') returns:
        ftplib.error_perm: 550 Failed to change directory.
        """
        raw_url = str(ftp_url).strip()
        parsed = urlparse(raw_url)
        host = parsed.hostname
        encoded_remote_path = parsed.path or ""
        remote_path = unquote(encoded_remote_path)
        url_user = unquote(parsed.username or "")
        if not host or not remote_path:
            raise ValueError(f"无效 FTP 地址：{ftp_url}")
        filename = os.path.basename(remote_path)
        remote_dir = os.path.dirname(remote_path) or "/"
        encoded_filename = os.path.basename(encoded_remote_path)
        encoded_remote_dir = os.path.dirname(encoded_remote_path) or "/"
        if not filename.lower().endswith(".zip"):
            raise ValueError(f"FTP 地址不是 .zip 文件：{ftp_url}")
        return host, remote_dir, filename, remote_path, url_user, encoded_remote_dir, encoded_filename, encoded_remote_path

    def get_default_output_folder(self):
        folder = os.path.join(os.path.expanduser("~/Downloads"), "OTADeployTool_Output")
        os.makedirs(folder, exist_ok=True)
        return folder

    def get_output_folder(self, source_file_path=""):
        if source_file_path:
            if os.path.isdir(source_file_path):
                return source_file_path
            return os.path.dirname(source_file_path)
        return self.get_default_output_folder()

    def create_temp_folder(self, source_file_path="", output_folder=None):
        base_folder = output_folder or self.get_output_folder(source_file_path)
        temp_folder = os.path.join(base_folder, "temp")
        os.makedirs(temp_folder, exist_ok=True)
        return temp_folder

    def download_ftp_file_binary_threaded(self, ftp_url, temp_folder, file_index, total_files):
        (
            host,
            remote_dir,
            filename,
            remote_path,
            url_user,
            encoded_remote_dir,
            encoded_filename,
            encoded_remote_path,
        ) = self.parse_ftp_url(ftp_url)
        local_path = os.path.join(temp_folder, filename)
        login_user = url_user or FTP_USER

        self.thread_log("")
        self.thread_log(f"下载任务 {file_index}/{total_files}")
        self.thread_log(f"原始 FTP URL：{ftp_url}")
        self.thread_log(f"FTP 服务器 IP/Host：{host}")
        self.thread_log(f"FTP 远程目录（解码后，用于 cwd）：{remote_dir}")
        self.thread_log(f"FTP 远程目录（URL 原始编码）：{encoded_remote_dir}")
        self.thread_log(f"ZIP 文件名（解码后）：{filename}")
        self.thread_log(f"ZIP 文件名（URL 原始编码）：{encoded_filename}")
        self.thread_log(f"FTP 远程完整路径（解码后）：{remote_path}")
        self.thread_log(f"FTP 远程完整路径（URL 原始编码）：{encoded_remote_path}")
        self.thread_log(f"本地暂存路径：{local_path}")
        self.thread_log(f"FTP 登录用户：{login_user}")
        self.thread_log("FTP 密码：已配置，日志中不打印明文。")

        ftp = None
        try:
            self.thread_log("正在建立 FTP 连接，超时时间：120 秒...")
            ftp = FTP(host, timeout=120)
            self.thread_log("FTP 连接建立成功，正在登录...")
            ftp.login(login_user, FTP_PASS)
            self.thread_log("FTP 登录服务器成功。")

            try:
                self.thread_log(f"FTP 登录后的当前目录：{ftp.pwd()}")
            except Exception:
                self.thread_log("FTP 当前目录读取失败。")

            cwd_success = False
            cwd_errors = []
            # Normal FTP path should be decoded.  Keep encoded path only as a
            # fallback because some non-standard servers may expose literal %xx.
            for candidate_dir, label in [
                (remote_dir, "解码后目录"),
                (encoded_remote_dir, "URL原始编码目录"),
            ]:
                if not candidate_dir:
                    continue
                try:
                    self.thread_log(f"正在切换到 FTP 文件目录（{label}）：{candidate_dir}")
                    ftp.cwd(candidate_dir)
                    self.thread_log(f"FTP 目录切换成功，当前目录：{ftp.pwd()}")
                    cwd_success = True
                    break
                except Exception as e:
                    cwd_errors.append(f"{label}={candidate_dir} -> {e}")
                    self.thread_log(f"FTP 目录切换失败（{label}）：{e}")

            if not cwd_success:
                self.thread_log("所有 FTP 目录切换方式均失败。失败详情：")
                for err in cwd_errors:
                    self.thread_log(f"  {err}")
                self.thread_log("将尝试使用完整远程路径直接 RETR 下载，避免依赖 cwd。")

            active_retr_name = filename if cwd_success else remote_path
            active_size_name = filename if cwd_success else remote_path

            try:
                if cwd_success:
                    names = ftp.nlst()
                    self.thread_log(f"当前目录文件数量：{len(names)}")
                    if filename in names:
                        self.thread_log("目标 ZIP 文件在当前目录中确认存在。")
                    else:
                        self.thread_log("警告：目录列表中未直接看到目标 ZIP 文件，仍将尝试 RETR 下载。")
                else:
                    self.thread_log("未进入目标目录，跳过目录列表读取。")
            except Exception as e:
                self.thread_log(f"FTP 目录列表读取失败，将继续尝试下载：{e}")

            try:
                file_size = ftp.size(active_size_name)
            except Exception:
                file_size = 0

            self.thread_log(f"准备下载 ZIP 文件：{filename}")
            self.thread_log(f"FTP RETR 使用路径：{active_retr_name}")
            if file_size:
                self.thread_log(f"文件大小：{file_size} bytes")
            else:
                self.thread_log("无法获取文件大小，将显示任务级进度。")

            downloaded = 0
            with open(local_path, "wb") as f:
                def callback(data):
                    nonlocal downloaded
                    f.write(data)
                    downloaded += len(data)
                    if file_size > 0:
                        file_percent = downloaded / file_size * 100
                        total_percent = ((file_index - 1) + file_percent / 100) / total_files * 100
                        self.thread_progress(total_percent, f"正在下载 {filename}：{file_percent:.1f}%")
                ftp.retrbinary(f"RETR {active_retr_name}", callback, blocksize=1024 * 128)

            self.thread_log(f"下载完成：{local_path}")
            self.thread_log(f"已下载字节数：{downloaded} bytes")
            total_percent = file_index / total_files * 100
            self.thread_progress(total_percent, f"已完成 {file_index}/{total_files} 个文件")
            return local_path
        finally:
            if ftp is not None:
                try:
                    ftp.quit()
                    self.thread_log("FTP 连接已关闭。")
                except Exception:
                    try:
                        ftp.close()
                    except Exception:
                        pass
                    self.thread_log("FTP 连接已强制关闭。")

    def scan_ftp_zip_urls_from_text(self, text):
        self.thread_log("开始扫描粘贴内容，查找 FTP ZIP 地址...")
        urls = [url for url in self.extract_zip_urls_from_text(text) if url.lower().startswith("ftp://")]
        self.thread_log(f"扫描完成：从粘贴内容中发现 FTP ZIP 地址 {len(urls)} 个。")
        return urls

    def download_all_ftp_files_worker(self, source_file_path="", device_code="", auto_upload=False, pasted_text=""):
        try:
            self.thread_log("开始后台下载 ZIP 文件...")
            if pasted_text:
                self.thread_log("数据来源：邮件/表格粘贴内容")
            else:
                self.thread_log(f"Source Excel 路径：{source_file_path}")
            self.thread_progress(0, "正在读取粘贴内容..." if pasted_text else "正在读取 Source Excel...")

            # First use normal field mapping; then scan the whole workbook as a fallback.
            # This guarantees that full FTP URLs in columns such as 整包地址 are found.
            try:
                source_json = self.build_source_json(source_file_path, pasted_text, use_gui_log=False)
                rows = source_json.get("rows", [])
                mapping = source_json.get("mapping", {})
                self.thread_log(f"字段映射识别结果：{self.format_mapping_for_log(mapping)}")
                self.thread_log(f"JSON OTA 数据行数：{len(rows)}")
                self.thread_log(f"JSON 升级提示语数量：{len(source_json.get('messages', []))}")
                mapped_urls = self.get_unique_ftp_urls(rows)
                self.thread_log(f"通过字段映射发现 FTP ZIP 文件数量：{len(mapped_urls)}")
            except Exception as e:
                mapped_urls = []
                self.thread_log(f"字段映射读取失败，将继续执行全表扫描：{e}")

            if pasted_text:
                scanned_urls = self.scan_ftp_zip_urls_from_text(pasted_text)
            else:
                scanned_urls = self.scan_ftp_zip_urls_from_excel(source_file_path)
            unique_urls = list(dict.fromkeys(mapped_urls + scanned_urls))

            self.thread_log("")
            self.thread_log("========== FTP 下载任务 ==========")
            self.thread_log(f"需要下载的 ZIP 文件总数：{len(unique_urls)}")

            if not unique_urls:
                message = "没有发现需要下载的 FTP ZIP 文件，请检查粘贴内容或 Source Excel 中的差分包地址。"
                self.thread_progress(100, message)
                self.thread_log(message)
                self.log_queue.put(("error", message))
                return

            self.thread_log("待下载文件列表：")
            for index, url in enumerate(unique_urls, start=1):
                self.thread_log(f"  {index}. {url}")

            temp_folder = self.create_temp_folder(source_file_path)
            self.thread_log(f"临时下载目录：{temp_folder}")

            downloaded_files = []
            failed_files = []
            for index, ftp_url in enumerate(unique_urls, start=1):
                try:
                    local_path = self.download_ftp_file_binary_threaded(ftp_url, temp_folder, index, len(unique_urls))
                    downloaded_files.append(local_path)
                    self.thread_log(f"下载成功 [{index}/{len(unique_urls)}]：{os.path.basename(local_path)}")
                except Exception as e:
                    failed_files.append((ftp_url, str(e)))
                    self.thread_log(f"下载失败 [{index}/{len(unique_urls)}]：{ftp_url}")
                    self.thread_log(f"失败原因：{e}")
                    self.thread_log(traceback.format_exc())

            self.thread_progress(100, "全部 ZIP 下载任务完成")
            self.thread_log("")
            self.thread_log("========== 下载完成 ==========")
            self.thread_log(f"需要下载 ZIP 文件总数：{len(unique_urls)}")
            self.thread_log(f"成功下载 ZIP 文件数：{len(downloaded_files)}")
            self.thread_log(f"失败下载 ZIP 文件数：{len(failed_files)}")
            if failed_files:
                self.thread_log("失败文件列表：")
                for ftp_url, reason in failed_files:
                    self.thread_log(f"  {ftp_url}")
                    self.thread_log(f"    {reason}")
                message = f"ZIP 下载存在失败项：成功 {len(downloaded_files)} 个，失败 {len(failed_files)} 个。已停止后续 AWS 上传动作。"
                self.thread_log(message)
                self.log_queue.put(("error", message))
                return

            if auto_upload and not device_code:
                message = "ZIP 已下载，但缺少设备类型代码，无法生成 AWS 上传路径，已停止上传。"
                self.thread_log(message)
                self.log_queue.put(("error", message))
                return

            if auto_upload and device_code and downloaded_files:
                temp_folder = self.create_temp_folder(source_file_path)
                self.thread_log("")
                self.thread_log("自动执行：ZIP 下载完成，开始 S3 上传...")
                self.upload_to_s3_worker(temp_folder, device_code)
            else:
                self.log_queue.put(("done", len(downloaded_files)))

        except Exception as e:
            self.thread_log("后台下载任务失败。")
            self.thread_log(str(e))
            self.thread_log(traceback.format_exc())
            self.log_queue.put(("error", str(e)))

    # ==========================================================
    # Data
    # ==========================================================

    def map_ota_type(self, value):
        """Normalize OTA upgrade type from the source Excel.

        Supported source values include:
        - 1 / 1.0 / 正常升级 / 1-正常升级 / 1 - 正常升级
        - 2 / 2.0 / 强制升级 / 2-强制升级 / 2 - 强制升级
        - 11 / 11.0 / FTE强制升级 / 11-FTE强制升级 / 11 - FTE强制升级

        The output template requires numeric codes only: 1, 2, or 11.
        """
        if value is None:
            return "1"

        raw = str(value).strip()
        if raw == "":
            return "1"

        # Normalize common punctuation and spacing variations from Excel cells.
        normalized = raw.replace("－", "-").replace("–", "-").replace("—", "-")
        normalized = normalized.replace("：", ":").replace("，", ",")
        normalized_no_space = re.sub(r"\s+", "", normalized)
        lowered = normalized_no_space.lower()

        # Numeric values may come from Excel as 1, 2, 11, 1.0, 2.0, 11.0.
        numeric_candidate = normalized_no_space
        if re.fullmatch(r"\d+(?:\.0+)?", numeric_candidate):
            numeric_code = str(int(float(numeric_candidate)))
            if numeric_code in ("1", "2", "11"):
                return numeric_code

        # Compound values like "11-FTE强制升级", "2 - 强制升级", "1:正常升级".
        prefix_match = re.match(r"^(11|2|1)(?:[-_:：、,，]|$)", normalized_no_space, re.IGNORECASE)
        if prefix_match:
            return prefix_match.group(1)

        # Text values. Check FTE before generic 强制, otherwise FTE强制升级 would map to 2.
        # Check 非强制升级 before generic 强制, otherwise the substring 强制 inside it
        # would cause a wrong match — 非强制升级 means normal upgrade and must map to 1.
        if "fte" in lowered and "强制" in normalized_no_space:
            return "11"
        if "非强制" in normalized_no_space:
            return "1"
        if "强制" in normalized_no_space:
            return "2"
        if "正常" in normalized_no_space:
            return "1"

        # English fallbacks.
        if "force" in lowered or "mandatory" in lowered:
            if "fte" in lowered:
                return "11"
            return "2"
        if "normal" in lowered or "regular" in lowered:
            return "1"

        # Keep previous behavior: default to normal upgrade. Caller logs raw value.
        return "1"

    def extract_zip_filename(self, url_or_path):
        if not url_or_path:
            return ""
        text = str(url_or_path).strip()
        match = re.search(r"([^/\\]+\.zip)(?:[?#].*)?$", text, re.IGNORECASE)
        if match:
            return unquote(match.group(1)).strip()
        parsed = urlparse(text)
        filename = os.path.basename(unquote(parsed.path or text))
        return filename.strip() if filename.lower().endswith(".zip") else ""

    def find_zip_source_url_in_row(self, row):
        """Find the FTP/HTTP/local ZIP source value for one source Excel row.

        Normal mapping should read 升级包地址, but some files name the column 整包地址
        or other variants.  This fallback scans every original cell in that row so the
        generated 升级文件地址 can always include the exact ZIP filename.
        """
        candidates = [
            self.get_value(row, "升级包地址"),
            self.get_value(row, "升级文件地址"),
            self.get_value(row, "整包地址"),
        ]
        candidates.extend(row.get("_all_cell_values", []) or [])
        for value in candidates:
            text = str(value or "").strip()
            if text and ".zip" in text.lower():
                filename = self.extract_zip_filename(text)
                if filename:
                    return text
        return ""

    def normalize_path_part(self, path):
        if not path:
            return ""
        return re.sub(r"/+", "/", str(path).strip().replace("\\", "/"))

    def sanitize_url(self, url):
        if not url:
            return url
        url = re.sub(r'^(https?:)/+', r'\1//', url, flags=re.IGNORECASE)
        scheme_end = url.find("://")
        if scheme_end != -1:
            scheme = url[:scheme_end + 3]
            rest = url[scheme_end + 3:]
            rest = re.sub(r'/+', '/', rest)
            url = scheme + rest
        else:
            url = re.sub(r'/+', '/', url)
        return url

    def build_upgrade_url(self, source_upgrade_url, aws_suffix):
        source_upgrade_url = source_upgrade_url or ""
        zip_filename = self.extract_zip_filename(source_upgrade_url)
        clean_suffix = self.normalize_path_part(aws_suffix or self.get_default_aws_suffix())
        if not clean_suffix.startswith("/"):
            clean_suffix = "/" + clean_suffix
        if not clean_suffix.endswith("/"):
            clean_suffix += "/"
        final_url = AWS_PREFIX.rstrip("/") + clean_suffix + zip_filename
        final_url = self.sanitize_url(final_url)
        return final_url, zip_filename

    def build_target_version(self, source_version, target_version, version_change):
        if version_change is not None and str(version_change).strip() != "":
            return str(version_change).strip(), False
        source_text = "" if source_version is None else str(source_version).strip()
        target_text = "" if target_version is None else str(target_version).strip()
        source_suffix = source_text[-5:] if len(source_text) >= 5 else source_text
        target_suffix = target_text[-5:] if len(target_text) >= 5 else target_text
        return f"{source_suffix}_to_{target_suffix}", True

    def safe_filename_part(self, text):
        text = str(text or "").strip()
        text = re.sub(r'[\\/:*?"<>|]+', "_", text)
        return text or "Device"

    def generate_output_path(self, source_file_path, device_code, row_count):
        folder = self.get_output_folder(source_file_path)
        device_part = self.safe_filename_part(device_code)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        return os.path.join(folder, f"{device_part}_{row_count}_{timestamp}.xlsx")

    def validate_output_row(self, row_index, row_data):
        warnings = []
        upgrade_url = row_data.get("升级文件地址", "")
        if upgrade_url:
            if not re.match(r'^https?://', upgrade_url, re.IGNORECASE):
                warnings.append(f"  ⚠ 升级文件地址 不是合法 URL（缺少 http/https 协议）：{upgrade_url}")
            if re.search(r'/{3,}', upgrade_url):
                warnings.append(f"  ⚠ 升级文件地址 包含非法连续斜杠（///）：{upgrade_url}")
            if not upgrade_url.lower().endswith(".zip"):
                warnings.append(f"  ⚠ 升级文件地址 不以 .zip 结尾：{upgrade_url}")
        else:
            warnings.append("  ⚠ 升级文件地址 为空")

        ota_type = str(row_data.get("OTA类型", "")).strip()
        if ota_type not in ("1", "2", "11"):
            warnings.append(f"  ⚠ OTA类型 值无效（期望 1 / 2 / 11，实际：'{ota_type}'）")

        if warnings:
            self.log_message(f"  [校验] 第 {row_index} 行发现问题：")
            for w in warnings:
                self.log_message(w)
        else:
            self.log_message(f"  [校验] 第 {row_index} 行校验通过 ✓")
        return warnings

    def get_value(self, row, field):
        return row.get(field, "")

    def format_mapping_for_log(self, mapping):
        if not mapping:
            return "未识别到任何字段"
        return ", ".join([f"{field}->{get_column_letter(col_idx)}" for field, col_idx in sorted(mapping.items(), key=lambda x: x[1])])

    def log_required_field_status(self, mapping):
        required_fields = ["源版本", "目标版本", "OTA类型", "升级包地址", "升级文件的MD5值", "文件大小"]
        self.log_message("关键字段检查：")
        for field in required_fields:
            if field in mapping:
                self.log_message(f"  ✓ 已识别：{field} -> {get_column_letter(mapping[field])}列")
            else:
                self.log_message(f"  ⚠ 未识别：{field}。生成时该字段将留空或使用默认规则。")
        if "区域" in mapping:
            self.log_message(f"  ✓ 已识别：区域 -> {get_column_letter(mapping['区域'])}列，将复制到输出文件的“区域组”。")
        else:
            self.log_message("  ⚠ 未识别：区域。输出文件“区域组”将留空。")
        if "MAC" in mapping:
            self.log_message(f"  说明：已检测到 MAC 列 {get_column_letter(mapping['MAC'])}。输出 MAC组优先使用 UI 输入；如果 UI 为空，将使用该列内容。")

    # ==========================================================
    # Workbook
    # ==========================================================

    def clean_file_size(self, value):
        if not value:
            return value
        return re.sub(r"[,\s_]", "", str(value).strip())

    def validate_feature_code(self, value):
        if not value:
            return value
        digits_only = re.sub(r"\D", "", str(value))
        if not digits_only:
            return value
        if len(digits_only) < 24:
            digits_only = digits_only.zfill(24)
        elif len(digits_only) > 24:
            digits_only = digits_only[-24:]
        return digits_only

    def apply_text_format(self, cell):
        cell.number_format = "@"
        cell.alignment = Alignment(vertical="center")

    def get_default_ota_messages(self):
        return [
            {"language": "简体中文", "description": "有新软件版本可用。更新包括：\n-改善用户体验\n-其他一些小问题\n\n请更新电视软件以改善用户体验。\n软件升级时，请勿断电。"},
            {"language": "英语", "description": "A new software version is available. Updates include:\n- Improve user experience\n- Some other minor fixes\n\nPlease update the TV software to improve your user experience.\nDo NOT turn off the power while the software is upgrading."},
            {"language": "西班牙语", "description": "Una nueva versión de software está disponible. Las actualizaciones incluyen:\n-Mejorar la experiencia del usuario\n-Algunas otras correcciones menores\n\nActualice el software de TV para mejorar su experiencia de usuario.\nCuando el S/W se está actualizando, NO apague la alimentación."},
            {"language": "繁体中文", "description": "有新軟體版本可用。更新包括：\n-改善用戶體驗\n-其他一些小問題\n\n請更新電視軟體以改善用戶體驗。\n軟體升級時，請勿斷電。"},
            {"language": "德语", "description": "Eine neue Softwareversion ist verfügbar. Aktualisierungen umfassen:\n- Verbesserung des Nutzungserlebnisses\n- weitere Bugfixes\n\nBitte aktualisieren Sie die TV-Software, um Ihre Benutzererfahrung zu verbessern.\nSchalten Sie die Stromversorgung NICHT aus, wenn die Software aktualisiert wird."},
            {"language": "葡萄牙语", "description": "Está disponível uma atualização de software importante.\nO que há de novo?\n-Correções de bugs de software\n-Melhorias de plataforma\n\nA alimentação tem de estar sempre ligada até à conclusão da atualização."},
        ]

    def order_ota_messages(self, messages):
        messages = list(messages or [])
        simplified = [item for item in messages if item.get("language") == "简体中文"]
        others = [item for item in messages if item.get("language") != "简体中文"]
        return simplified + others

    def ensure_simplified_chinese_first(self, messages):
        """Output rule for the generated "升级提示语" sheet.

        简体中文 must always be the first data row. When the source provides no
        Chinese prompt at all, duplicate the English entry as a 简体中文 row so the
        first row is still 简体中文. If neither 简体中文 nor 英语 exists, the list is
        returned unchanged (only reordered so any existing 简体中文 stays first).
        """
        messages = [dict(item) for item in (messages or []) if item]
        if any(item.get("language") == "简体中文" for item in messages):
            return self.order_ota_messages(messages)
        english = next((item for item in messages if item.get("language") == "英语"), None)
        if english is None:
            return self.order_ota_messages(messages)
        clone = {"language": "简体中文", "description": english.get("description", "")}
        return [clone] + messages

    def standardize_ota_messages(self, ota_messages):
        messages = []
        for item in ota_messages or []:
            if isinstance(item, dict):
                language = item.get("language", "")
                description = item.get("description", "")
            else:
                language = item[0] if len(item) > 0 else ""
                description = item[1] if len(item) > 1 else ""
            language = self.normalize_message_language(language)
            description = str(description or "").strip()
            if language and description:
                messages.append({"language": language, "description": description})
        return self.order_ota_messages(self.dedupe_language_messages(messages))

    def create_workbook(self, output_rows, ota_messages=None):
        wb = Workbook()
        ws = wb.active
        ws.title = "升级规则"

        header_fill = PatternFill("solid", fgColor="D9EAF7")
        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for col_idx, header in enumerate(OUTPUT_HEADERS, start=1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
            cell.number_format = "@"

        for row_idx, row_data in enumerate(output_rows, start=2):
            for col_idx, header in enumerate(OUTPUT_HEADERS, start=1):
                cell = ws.cell(row=row_idx, column=col_idx, value=row_data.get(header, ""))
                self.apply_text_format(cell)
                cell.border = border

        ws.freeze_panes = "A2"

        # ===============================
        # OTA升级提示语 Sheet
        # ===============================
        msg_ws = wb.create_sheet("升级提示语")

        msg_headers = ["语种", "描述"]
        for col_idx, header in enumerate(msg_headers, start=1):
            cell = msg_ws.cell(row=1, column=col_idx, value=header)
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border

        message_rows = self.standardize_ota_messages(ota_messages) if ota_messages else self.get_default_ota_messages()
        message_rows = self.ensure_simplified_chinese_first(message_rows)

        for row_idx, row in enumerate(message_rows, start=2):
            values = [row.get("language", ""), row.get("description", "")] if isinstance(row, dict) else row
            for col_idx, value in enumerate(values, start=1):
                cell = msg_ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = border
                cell.alignment = Alignment(wrap_text=True, vertical="top")

        msg_ws.column_dimensions["A"].width = 20
        msg_ws.column_dimensions["B"].width = 120


        for col_idx, header in enumerate(OUTPUT_HEADERS, start=1):
            max_len = len(header)
            for row_idx in range(2, ws.max_row + 1):
                value = ws.cell(row=row_idx, column=col_idx).value
                if value is not None:
                    max_len = max(max_len, len(str(value)))
            ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 4, 90)
        return wb

    # ==========================================================
    # Rule ID columns
    # ==========================================================

    def parse_rule_id_range(self, raw_value):
        """Parse a Rule ID input into an ordered list of integer IDs.

        Accepts a single id (``65213``) or an inclusive range (``65210-65213``).
        Raises ValueError with a user-facing message on invalid input.
        """
        text = str(raw_value or "").strip()
        if not text:
            raise ValueError("请先输入 Rule ID，例如 65213 或 65210-65213。")
        # Normalize common separators (full-width dash, en/em dash, "~", "到").
        normalized = re.sub(r"[－–—~～]|到", "-", text).replace(" ", "")

        def require_5_digits(token):
            if len(token) != 5:
                raise ValueError(f"Rule ID 必须是 5 位数：{token} 为 {len(token)} 位。")

        if re.fullmatch(r"\d+", normalized):
            require_5_digits(normalized)
            return [int(normalized)]
        range_match = re.fullmatch(r"(\d+)-(\d+)", normalized)
        if range_match:
            start_token, end_token = range_match.group(1), range_match.group(2)
            require_5_digits(start_token)
            require_5_digits(end_token)
            start, end = int(start_token), int(end_token)
            if end < start:
                raise ValueError(f"Rule ID 范围无效：结束值 {end} 小于起始值 {start}。")
            ids = list(range(start, end + 1))
            # A range could cross into 6 digits (e.g. 99998-100002); every id must
            # stay 5 digits.
            for rid in ids:
                if len(str(rid)) != 5:
                    raise ValueError(f"Rule ID 必须是 5 位数：范围内的 {rid} 为 {len(str(rid))} 位。")
            return ids
        raise ValueError("Rule ID 格式不正确。请输入单个数字（如 65213）或范围（如 65210-65213）。")

    def insert_rule_id_columns(self, file_path, rule_ids, sheet_name="升级规则", auto_increment=False):
        """Insert 升级方式 / 升级日期 / Rule ID columns into a generated workbook.

        The three columns are inserted immediately after the "OTA类型" column,
        preserving the order and content of every other column. 升级方式 is derived
        from the existing OTA类型 code, 升级日期 is today's date (YYYYMMDD) and
        Rule ID is assigned per data row. Returns (data_row_count, assigned_ids).
        """
        if not file_path or not os.path.exists(file_path):
            raise ValueError("要更新的 Excel 文件不存在，请先生成模版文件。")

        wb = load_workbook(file_path)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[0]

        headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]

        # Idempotency guard: refuse to insert a second time.
        if any(h in RULE_ID_COLUMNS for h in headers):
            raise ValueError("该文件已包含 Rule ID 相关列，无需重复插入。")

        try:
            ota_type_pos = headers.index(RULE_ID_INSERT_AFTER_HEADER) + 1  # 1-based
        except ValueError:
            raise ValueError(f"未在表头中找到“{RULE_ID_INSERT_AFTER_HEADER}”列，无法定位插入位置。")

        ota_col = ota_type_pos  # the OTA类型 column index
        insert_at = ota_type_pos + 1  # insert the 3 columns right after it

        data_row_count = 0
        for r in range(2, ws.max_row + 1):
            if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, ws.max_column + 1)):
                data_row_count = r - 1  # last non-empty data row index (row 2 == row #1)

        # Rule ID assignment:
        # - auto_increment (a single start id was given): fill每行 by +1 from the
        #   start id, covering all data rows automatically.
        # - otherwise (an explicit range/list): the count must match the rows.
        if auto_increment and len(rule_ids) == 1 and data_row_count > 0:
            start = rule_ids[0]
            assigned_ids = [start + i for i in range(data_row_count)]
            over = [rid for rid in assigned_ids if len(str(rid)) != 5]
            if over:
                raise ValueError(
                    f"自动累加后 Rule ID 超出 5 位数（如 {over[0]}）。请换一个更小的起始 Rule ID。"
                )
        else:
            if len(rule_ids) != data_row_count:
                raise ValueError(
                    f"Rule ID 数量与数据行数不匹配：提供了 {len(rule_ids)} 个 Rule ID，"
                    f"但 Excel 中有 {data_row_count} 行数据。请提供正好 {data_row_count} 个 Rule ID，"
                    f"或只输入起始 Rule ID 让系统自动累加。"
                )
            assigned_ids = list(rule_ids)
        today = datetime.now().strftime("%Y%m%d")

        # openpyxl's insert_cols shifts cell data but does NOT shift the existing
        # column_dimensions (widths). Capture every original width by column index
        # BEFORE inserting so we can re-map them to the correct columns afterwards,
        # leaving the rest of the layout exactly as it was generated.
        orig_col_count = ws.max_column
        orig_widths = {}
        for c in range(1, orig_col_count + 1):
            dim = ws.column_dimensions.get(get_column_letter(c))
            if dim is not None and dim.width is not None:
                orig_widths[c] = dim.width

        ws.insert_cols(insert_at, amount=len(RULE_ID_COLUMNS))

        # Re-apply the captured widths: columns before the insertion point keep
        # their letter; columns at/after it shift right by len(RULE_ID_COLUMNS).
        shift_amount = len(RULE_ID_COLUMNS)
        for c in range(orig_col_count, 0, -1):
            new_c = c + shift_amount if c >= insert_at else c
            if c in orig_widths:
                ws.column_dimensions[get_column_letter(new_c)].width = orig_widths[c]
            elif new_c != c:
                # Clear any width that leaked onto the shifted slot.
                ws.column_dimensions[get_column_letter(new_c)].width = None

        # Carry over styling from a reference cell so the new columns match.
        header_fill = PatternFill("solid", fgColor="D9EAF7")
        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        for offset, name in enumerate(RULE_ID_COLUMNS):
            col = insert_at + offset
            cell = ws.cell(row=1, column=col, value=name)
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = border
            cell.number_format = "@"

        # Locate columns needed for the deployment summary. `headers` is the
        # PRE-insertion layout; any column at or after the insertion point shifted
        # right by len(RULE_ID_COLUMNS) once the new columns were inserted.
        shift = len(RULE_ID_COLUMNS)

        def header_col(name):
            if name not in headers:
                return None
            pos = headers.index(name) + 1  # 1-based, pre-insertion
            return pos + shift if pos >= insert_at else pos

        device_code_col = header_col("设备类型代码")
        model_col = header_col("内部机型信息")
        ext_col = header_col("设备扩展信息")
        source_ver_col = header_col("源版本")
        target_ver_col = header_col("目标版本")

        def cell_text(r, col):
            return str(ws.cell(row=r, column=col).value or "").strip() if col else ""

        deployment_records = []
        for r in range(2, data_row_count + 2):
            ota_code = str(ws.cell(row=r, column=ota_col).value or "").strip()
            upgrade_method = OTA_TYPE_CODE_TO_TEXT.get(ota_code, "")
            rule_id = assigned_ids[r - 2] if (r - 2) < len(assigned_ids) else ""
            values = [upgrade_method, today, str(rule_id)]
            for offset, value in enumerate(values):
                cell = ws.cell(row=r, column=insert_at + offset, value=value)
                self.apply_text_format(cell)
                cell.border = border
            deployment_records.append({
                "设备类型代码": cell_text(r, device_code_col),
                "Rule ID": str(rule_id),
                "机型信息": cell_text(r, model_col),
                "扩展码": cell_text(r, ext_col),
                "源版本": cell_text(r, source_ver_col),
                "目标版本": cell_text(r, target_ver_col),
                "部署时间": today,
            })

        # Width for the three new columns.
        widths = {RULE_ID_COLUMNS[0]: 16, RULE_ID_COLUMNS[1]: 14, RULE_ID_COLUMNS[2]: 12}
        for offset, name in enumerate(RULE_ID_COLUMNS):
            ws.column_dimensions[get_column_letter(insert_at + offset)].width = widths[name]

        try:
            wb.save(file_path)
        except PermissionError:
            raise ValueError(
                "保存失败：文件可能正在被 Excel / Numbers 打开并锁定。\n"
                "请关闭已打开的该文件后重试。"
            )
        return data_row_count, assigned_ids, deployment_records

    def read_excel_preview(self, file_path, sheet_name="升级规则", max_rows=200):
        """Read a workbook's first (升级规则) sheet into {headers, rows} for preview.

        Returns the header row and up to `max_rows` non-empty data rows so the UI
        can show the current on-disk content of the generated file (including the
        just-inserted Rule ID columns) without forcing a download.
        """
        if not file_path or not os.path.exists(file_path):
            raise ValueError("要预览的 Excel 文件不存在，请先生成模版文件。")

        wb = load_workbook(file_path, data_only=True)
        ws = wb[sheet_name] if sheet_name in wb.sheetnames else wb.worksheets[0]

        headers = [
            str(ws.cell(row=1, column=c).value or "").strip()
            for c in range(1, ws.max_column + 1)
        ]

        rows = []
        for r in range(2, ws.max_row + 1):
            values = [
                str(ws.cell(row=r, column=c).value if ws.cell(row=r, column=c).value is not None else "")
                for c in range(1, ws.max_column + 1)
            ]
            if any(v.strip() for v in values):
                rows.append(values)
            if len(rows) >= max_rows:
                break

        return {
            "sheet": ws.title,
            "headers": headers,
            "rows": rows,
            "row_count": len(rows),
            "truncated": len(rows) >= max_rows,
        }

    def _display_width(self, text):
        """Visual width where CJK characters count as 2 columns."""
        width = 0
        for ch in str(text):
            width += 2 if ord(ch) > 0x2E7F else 1
        return width

    def _pad_cell(self, text, width):
        text = str(text)
        pad = width - self._display_width(text)
        return text + (" " * pad if pad > 0 else "")

    def build_rule_id_deployment_summary(self, deployment_records, file_path=""):
        """Build a copyable, boxed-table summary of the OTA deployment.

        Columns: Rule ID, 机型信息, 机型扩展码, 源版本, 目标版本, 部署时间. The table is
        drawn with box characters, and ends with a reminder to carry the Rule ID
        info into the next gray release.
        """
        columns = ["Rule ID", "机型信息", "扩展码", "源版本", "目标版本", "部署时间"]
        display_headers = {"扩展码": "机型扩展码"}

        # Column width = max display width of header and all values, +1 breathing room.
        widths = {}
        for col in columns:
            w = self._display_width(display_headers.get(col, col))
            for rec in deployment_records:
                w = max(w, self._display_width(rec.get(col, "")))
            widths[col] = w + 1

        def border(left, mid, right):
            return left + mid.join("─" * (widths[col] + 2) for col in columns) + right

        def render_row(cells):
            return "│ " + " │ ".join(self._pad_cell(cells.get(col, ""), widths[col]) for col in columns) + " │"

        all_ids = [rec["Rule ID"] for rec in deployment_records if str(rec.get("Rule ID")).strip()]
        rule_id_span = ""
        if all_ids:
            rule_id_span = all_ids[0] if len(all_ids) == 1 else f"{all_ids[0]} ~ {all_ids[-1]}"

        header_cells = {col: display_headers.get(col, col) for col in columns}

        # Device type code is one value for the whole deployment; show it once on
        # the summary (总数) line rather than repeating it in every table row.
        device_code = ""
        for rec in deployment_records:
            dc = str(rec.get("设备类型代码", "") or "").strip()
            if dc:
                device_code = dc
                break

        lines = []
        lines.append("📦 OTA 部署信息汇总 / OTA Deployment Summary")
        meta = f"   部署规则总数：{len(deployment_records)} 条"
        if device_code:
            meta = f"   设备类型代码：{device_code}    " + meta.strip()
        if rule_id_span:
            meta += f"    Rule ID 范围：{rule_id_span}"
        lines.append(meta)
        lines.append("")
        lines.append(border("┌", "┬", "┐"))
        lines.append(render_row(header_cells))
        lines.append(border("├", "┼", "┤"))
        for rec in deployment_records:
            lines.append(render_row(rec))
        lines.append(border("└", "┴", "┘"))

        lines.append("")
        lines.append("⚠  灰度提醒：进行下一阶段灰度发布时，请提供本次的 Rule ID 信息。")
        if rule_id_span:
            lines.append(f"   本次部署 Rule ID 范围：{rule_id_span}")
        return "\n".join(lines)

    def resolve_generated_excel_path(self):
        """Find the Excel file the Rule ID update should operate on.

        Prefers the most recently generated output recorded during generation;
        otherwise, if the first field points directly at an .xlsx file, use it.
        """
        last_output = getattr(self, "last_generated_output_path", "")
        if last_output and os.path.exists(last_output):
            return last_output
        source_file = self.source_file_var.get().strip()
        if source_file and os.path.isfile(source_file) and source_file.lower().endswith((".xlsx", ".xlsm")):
            return source_file
        return ""

    def update_rule_id_columns(self):
        try:
            rule_ids = self.parse_rule_id_range(self.rule_id_range_var.get())
        except ValueError as e:
            self.show_center_error("Rule ID 输入错误", str(e))
            return

        target_path = self.resolve_generated_excel_path()
        if not target_path:
            target_path = filedialog.askopenfilename(
                title="选择要更新的已生成 Excel 文件",
                filetypes=[("Excel 文件", "*.xlsx *.xlsm")],
            )
        if not target_path:
            self.show_center_error("未选择文件", "请先生成模版文件，或选择要更新的 Excel 文件。")
            return

        try:
            row_count, assigned_ids, deployment_records = self.insert_rule_id_columns(target_path, rule_ids)
        except Exception as e:
            self.log_message(f"更新 Rule ID 列失败：{e}")
            self.show_center_error("更新失败", str(e))
            return

        id_desc = (
            f"{assigned_ids[0]}~{assigned_ids[-1]}" if assigned_ids and len(assigned_ids) > 1
            else (str(assigned_ids[0]) if assigned_ids else "无")
        )
        self.log_message(
            f"已在“{RULE_ID_INSERT_AFTER_HEADER}”列后插入 {('、'.join(RULE_ID_COLUMNS))}；"
            f"数据行数：{row_count}，Rule ID：{id_desc}，升级日期：{datetime.now().strftime('%Y%m%d')}。"
        )

        self.log_message(f"已在后台更新文件：{target_path}")
        self.log_message("如果该 Excel 已在 Excel / Numbers 中打开，请重新打开或刷新以查看更新。")

        # Print the copyable deployment summary into the log (no popup).
        summary = self.build_rule_id_deployment_summary(deployment_records, file_path=target_path)
        self.log_message("")
        self.log_message(summary)
        self.log_message("")
        self.log_message("提示：可点击下方“复制日志 / Copy Log”复制完整部署汇总。")

    # ==========================================================
    # Validation
    # ==========================================================

    def validate_inputs(self, source_file_path, pasted_text=""):
        if pasted_text:
            if source_file_path:
                if not os.path.exists(source_file_path):
                    raise ValueError("指定的工作目录或 Source Excel 文件不存在。")
                if not os.path.isdir(source_file_path) and not source_file_path.lower().endswith((".xlsx", ".xlsm")):
                    raise ValueError("粘贴内容模式下，第一行请输入工作目录，或 .xlsx / .xlsm 文件路径。")
            return
        if not source_file_path:
            raise ValueError("请先粘贴邮件/表格内容，或选择 Source Excel 文件。")
        if not os.path.exists(source_file_path):
            raise ValueError("Source Excel 文件不存在。")
        if os.path.isdir(source_file_path):
            raise ValueError("未粘贴内容时，请选择具体的 Source Excel 文件，而不是目录。")
        if not source_file_path.lower().endswith((".xlsx", ".xlsm")):
            raise ValueError("请选择 .xlsx 或 .xlsm 格式的 Excel 文件。")

    def prepare_common(self, animation_text):
        self.log_message("")
        self.log_message("=" * 60)
        self.log_message(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {animation_text}")
        self.log_message("=" * 60)

        self.download_progress_var.set(0)
        self.download_status_var.set("下载进度：0%")

        source_file_path = self.source_file_var.get().strip()
        pasted_text = self.get_pasted_source_text()
        source_kind = "粘贴内容" if pasted_text else "Source Excel"
        self.log_message(f"数据来源：{source_kind}")
        if pasted_text:
            self.log_message(f"粘贴内容长度：{len(pasted_text)} 字符")
            if source_file_path:
                self.log_message(f"输出工作目录/参考文件：{source_file_path}")
            self.apply_pasted_form_fields(pasted_text, show_log=True)
        else:
            self.log_message(f"Source Excel 路径：{source_file_path}")
            self.clear_auto_fill_highlights()
        self.validate_inputs(source_file_path, pasted_text)
        self.log_message("输入文件校验通过。")
        self.update_aws_suffix_from_device_code()
        self.log_message(f"设备类型代码：{self.device_code_var.get().strip() or '(空)'}")
        self.log_message(f"设备类型名称：{self.device_name_var.get().strip() or '(空)'}")
        self.log_message(f"Feature Code / 特征码：{self.feature_code_var.get().strip() or '(空)'}")
        self.log_message(f"MAC 地址组：{self.mac_group_var.get().strip() or '(空)'}")
        self.log_message(f"AWS 下载地址前缀：{AWS_PREFIX}")
        self.log_message(f"AWS 文件路径后缀：{self.aws_suffix_var.get().strip()}")
        self.log_message(f"AWS UI 完整显示路径：{self.aws_full_location_var.get().strip()}")
        self.save_config()
        self.log_message(f"配置已保存：{CONFIG_FILE}")
        self.disable_all_buttons()
        self.start_animation(animation_text)
        return source_file_path, pasted_text

    # ==========================================================
    # Buttons
    # ==========================================================

    def generate_template_only(self):
        self.generate_excel_template()

    def download_upload_zip_package(self):
        try:
            source_file_path, pasted_text = self.prepare_common("正在下载/上传ZIP升级包")
            device_code_for_zip = self.device_code_var.get().strip()
            if not device_code_for_zip:
                raise ValueError("请先填写或粘贴可识别的设备类型代码，AWS 上传路径需要该字段。")
            self.log_message("按钮：下载/上传ZIP升级包")
            self.log_message("ZIP 下载 + AWS 上传任务已启动，正在后台运行。")
            self.log_message("只有全部 ZIP 下载成功后，系统才会继续上传到 AWS。")
            self.log_message("界面不会卡住，可以继续查看日志和进度。")
            worker = threading.Thread(
                target=self.download_all_ftp_files_worker,
                args=(source_file_path, device_code_for_zip, True, pasted_text),
                daemon=True
            )
            worker.start()
        except Exception as e:
            self.stop_animation("下载/上传失败")
            self.log_message(str(e))
            self.log_message(traceback.format_exc())
            self.show_center_error("错误", str(e))

    def download_zip_only(self):
        self.download_upload_zip_package()

    def auto_execute_all(self):
        self.generate_excel_template(start_zip_after=True)

    def upload_to_s3_only(self):
        try:
            source_file_path, pasted_text = self.prepare_common("正在上传至S3")
            device_code = self.device_code_var.get().strip()
            if not device_code:
                raise ValueError("请先填写 设备类型代码，S3 路径需要该字段。")
            temp_folder = self.create_temp_folder(source_file_path)
            self.log_message("按钮：上传ZIP文件到AWS")
            self.log_message("S3 上传任务已启动，正在后台运行。")
            self.log_message("界面不会卡住，可以继续查看日志和进度。")
            worker = threading.Thread(target=self.upload_to_s3_worker, args=(temp_folder, device_code), daemon=True)
            worker.start()
        except Exception as e:
            self.stop_animation("上传失败")
            self.log_message(str(e))
            self.log_message(traceback.format_exc())
            self.show_center_error("错误", str(e))

    def check_aws_cli_ready(self):
        try:
            aws_path = shutil.which("aws")
            if aws_path is None:
                self.thread_log("AWS CLI is not installed or not found in PATH.")
                self.thread_log("请先安装 AWS CLI，并确认 Terminal 中可以执行：aws --version")
                return False

            self.thread_log(f"AWS CLI 路径：{aws_path}")
            version_result = subprocess.run(["aws", "--version"], capture_output=True, text=True, timeout=20)
            version_text = (version_result.stdout or version_result.stderr).strip()
            if version_text:
                self.thread_log(f"AWS CLI 版本：{version_text}")
            result = subprocess.run(["aws", "sts", "get-caller-identity"], capture_output=True, text=True, timeout=20)
            if result.returncode == 0:
                self.thread_log("AWS CLI is configured and ready.")
                identity = result.stdout.strip()
                if identity:
                    self.thread_log(f"AWS 当前身份：{identity}")
                return True

            self.thread_log("AWS CLI is not configured correctly. Please run `aws configure` first.")
            if result.stderr.strip():
                self.thread_log(f"AWS 错误信息：{result.stderr.strip()}")
            elif result.stdout.strip():
                self.thread_log(f"AWS 输出信息：{result.stdout.strip()}")
            return False

        except subprocess.TimeoutExpired:
            self.thread_log("AWS CLI 检查超时，请检查网络、AWS CLI 配置或权限。")
            return False
        except Exception as e:
            self.thread_log("AWS CLI 检查异常，无法继续执行 S3 操作。")
            self.thread_log(str(e))
            return False

    # ==========================================================
    # S3 Upload
    # ==========================================================

    def build_s3_destination(self, device_code):
        """Return the S3 destination prefix for today's upload.
        Format: s3://fam-media-andr/ota/{device_code}/{YYYYMMDD}/
        """
        today = datetime.now().strftime("%Y%m%d")
        device_safe = device_code.strip() or "unknown"
        return f"s3://fam-media-andr/ota/{device_safe}/{today}/"

    def cleanup_uploaded_zip_files(self, temp_folder):
        """Remove local ZIP artifacts after all S3 uploads succeeded."""
        base_folder = os.path.dirname(os.path.abspath(temp_folder))
        temp_folder = os.path.abspath(temp_folder)

        if os.path.basename(temp_folder) != "temp":
            self.thread_log(f"安全检查：临时目录名称不是 temp，跳过删除目录：{temp_folder}")
            return

        self.thread_log("")
        self.thread_log("========== 本地 ZIP 清理 ==========")
        self.thread_log(f"当前目录：{base_folder}")
        self.thread_log(f"临时目录：{temp_folder}")

        removed_zip_count = 0
        failed_cleanup = []

        if os.path.isdir(base_folder):
            for filename in os.listdir(base_folder):
                local_path = os.path.join(base_folder, filename)
                if os.path.isfile(local_path) and filename.lower().endswith(".zip"):
                    try:
                        os.remove(local_path)
                        removed_zip_count += 1
                        self.thread_log(f"已删除当前目录 ZIP：{local_path}")
                    except Exception as e:
                        failed_cleanup.append((local_path, str(e)))
                        self.thread_log(f"删除当前目录 ZIP 失败：{local_path}，原因：{e}")

        if os.path.isdir(temp_folder):
            try:
                shutil.rmtree(temp_folder)
                self.thread_log(f"已删除 temp 文件夹：{temp_folder}")
            except Exception as e:
                failed_cleanup.append((temp_folder, str(e)))
                self.thread_log(f"删除 temp 文件夹失败：{temp_folder}，原因：{e}")
        else:
            self.thread_log("temp 文件夹不存在，无需删除。")

        self.thread_log(f"已删除当前目录 ZIP 文件数量：{removed_zip_count}")
        if failed_cleanup:
            self.thread_log(f"本地清理有 {len(failed_cleanup)} 项失败，请手动检查。")
        else:
            self.thread_log("本地 ZIP 清理完成。")

    def upload_to_s3_worker(self, temp_folder, device_code):
        try:
            self.thread_log("")
            self.thread_log("========== S3 上传任务 ==========")
            self.thread_log("正在检查 AWS CLI 安装和配置状态...")

            if not self.check_aws_cli_ready():
                self.thread_log("S3 上传已停止：AWS CLI 未准备好。")
                self.log_queue.put(("error", "S3 上传已停止：AWS CLI 未安装、未配置或当前账号无权限。请查看执行日志。"))
                return

            if not os.path.isdir(temp_folder):
                self.thread_log(f"临时目录不存在，请先执行 下载ZIP文件：{temp_folder}")
                self.log_queue.put(("error", f"S3 上传已停止：临时目录不存在：{temp_folder}"))
                return

            zip_files = [os.path.join(temp_folder, f) for f in os.listdir(temp_folder) if f.lower().endswith(".zip")]
            if not zip_files:
                self.thread_log(f"临时目录中未找到任何 .zip 文件：{temp_folder}")
                self.log_queue.put(("error", f"S3 上传已停止：临时目录中未找到 ZIP 文件：{temp_folder}"))
                return

            s3_dest = self.build_s3_destination(device_code)
            self.thread_log(f"找到 {len(zip_files)} 个 ZIP 文件")
            self.thread_log(f"S3 目标路径：{s3_dest}")
            self.thread_log("")

            success_count = 0
            fail_count = 0
            total = len(zip_files)

            for idx, local_path in enumerate(zip_files, start=1):
                filename = os.path.basename(local_path)
                s3_target = s3_dest + filename
                self.thread_log(f"[{idx}/{total}] 正在上传：{filename}")
                self.thread_log(f"  本地：{local_path}")
                self.thread_log(f"  目标：{s3_target}")
                self.thread_progress((idx - 1) / total * 100, f"正在上传 {filename}（{idx}/{total}）")

                try:
                    cmd = ["aws", "s3", "cp", local_path, s3_target, "--no-progress"]
                    result = subprocess.run(cmd, capture_output=True, text=True, timeout=600)
                    if result.returncode == 0:
                        self.thread_log(f"  ✓ 上传成功：{filename}")
                        if result.stdout.strip():
                            self.thread_log(f"  {result.stdout.strip()}")
                        success_count += 1
                    else:
                        self.thread_log(f"  ✗ 上传失败：{filename}")
                        if result.stderr.strip():
                            self.thread_log(f"  错误信息：{result.stderr.strip()}")
                        fail_count += 1
                except subprocess.TimeoutExpired:
                    self.thread_log(f"  ✗ 上传超时（10分钟）：{filename}")
                    fail_count += 1
                except FileNotFoundError:
                    self.thread_log("  ✗ 未找到 aws 命令，请确认已安装 AWS CLI 并配置好 PATH。")
                    self.thread_log("  安装参考：https://docs.aws.amazon.com/cli/latest/userguide/install-cliv2.html")
                    fail_count += 1
                    break
                except Exception as e:
                    self.thread_log(f"  ✗ 上传异常：{e}")
                    fail_count += 1

            self.thread_progress(100, f"S3 上传完成：成功 {success_count} / 失败 {fail_count}")
            self.thread_log("")
            self.thread_log("========== S3 上传结果 ==========")
            self.thread_log(f"S3 目标路径：{s3_dest}")
            self.thread_log(f"成功上传：{success_count} 个文件")
            if fail_count:
                self.thread_log(f"上传失败：{fail_count} 个文件")
            elif success_count == total:
                self.thread_log("所有 ZIP 文件均已成功上传，开始清理本地 ZIP 文件和 temp 文件夹。")
                self.cleanup_uploaded_zip_files(temp_folder)
            self.log_queue.put(("s3_done", success_count, fail_count))

        except Exception as e:
            self.thread_log("S3 上传任务异常终止。")
            self.thread_log(str(e))
            self.thread_log(traceback.format_exc())
            self.log_queue.put(("error", str(e)))

    # ==========================================================
    # Generate Excel
    # ==========================================================

    def confirm_mac_group_before_generate(self, mac_group=None):
        mac_group = str(mac_group if mac_group is not None else self.mac_group_var.get()).strip() or "(空)"
        return messagebox.askyesno(
            "确认 MAC 地址组",
            f"请确认 MAC 地址组是否正确：\n\n{mac_group}\n\n是否继续生成 Excel 模版？",
            parent=self.root,
        )

    def generate_excel_template(self, start_zip_after=False):
        try:
            source_file_path, pasted_text = self.prepare_common("正在自动执行" if start_zip_after else "正在生成模版文件")
            device_code = self.device_code_var.get().strip()
            device_name = self.device_name_var.get().strip()
            feature_code = self.feature_code_var.get().strip()
            aws_suffix = self.get_default_aws_suffix(device_code)
            self.aws_suffix_var.set(aws_suffix)
            self.aws_full_location_var.set(AWS_PREFIX.rstrip("/") + aws_suffix)
            mac_group = self.mac_group_var.get().strip()

            self.log_message("按钮：自动执行" if start_zip_after else "按钮：生成模版文件")
            self.log_message("开始生成 OTA 部署文件...")
            output_folder = self.get_output_folder(source_file_path)
            self.log_message(f"输出文件将保存到：{output_folder}")

            source_json = self.build_source_json(source_file_path, pasted_text, use_gui_log=True)
            rows = source_json.get("rows", [])
            mapping = source_json.get("mapping", {})
            source_ota_messages = source_json.get("messages", [])
            ota_messages, ota_message_source = self.select_ota_messages_for_generation(
                source_ota_messages,
                device_code,
                source_type=source_json.get("source_type", ""),
                prompt_user=True,
            )
            self.log_message(f"是否检测到“版本变化”列：{'是' if '版本变化' in mapping else '否'}")
            self.log_message(f"是否检测到“区域”列：{'是，将复制到输出“区域组”' if '区域' in mapping else '否，输出“区域组”将留空'}")
            self.log_message(f"是否检测到 MAC 列：{'是，UI为空时将用于输出MAC组' if 'MAC' in mapping else '否'}")
            self.log_message(f"是否检测到自定义升级提示语：{'是，共 ' + str(len(source_ota_messages)) + ' 条' if source_ota_messages else '否'}")
            if ota_message_source == "source_saved":
                self.log_message(f"升级提示语来源：本次粘贴内容，共 {len(ota_messages)} 条；已保存为设备类型代码 {device_code} 的参考模板。")
            elif ota_message_source == "source":
                self.log_message(f"升级提示语来源：源数据自带提示语，共 {len(ota_messages)} 条。")
            elif ota_message_source == "device_template_selected":
                self.log_message(f"升级提示语来源：用户选择的已有模板，共 {len(ota_messages)} 条。")
            elif ota_message_source == "device_template":
                self.log_message(f"升级提示语来源：设备类型代码 {device_code} 的内置参考模板，共 {len(ota_messages)} 条。")
            else:
                self.log_message(f"升级提示语来源：通用默认提示语，共 {len(ota_messages)} 条。")
            source_mac_group = next((self.get_value(row, "MAC") for row in rows if self.get_value(row, "MAC")), "")
            effective_mac_group = mac_group or source_mac_group
            if source_mac_group and not mac_group:
                self.log_message(f"检测到粘贴/源数据中的 MAC 地址组，将用于输出：{source_mac_group}")
                self.mac_group_var.set(source_mac_group)
            if not self.confirm_mac_group_before_generate(effective_mac_group):
                self.stop_animation("已取消生成")
                self.log_message("用户取消生成：MAC 地址组未确认。")
                return
            output_rows = []
            auto_target_version_count = 0

            for index, row in enumerate(rows, start=1):
                source_version = self.get_value(row, "源版本")
                full_target_version = self.get_value(row, "目标版本")
                version_change = self.get_value(row, "版本变化")
                ota_type_raw = self.get_value(row, "OTA类型")
                mapped_upgrade_url = self.get_value(row, "升级包地址")
                source_upgrade_url = self.find_zip_source_url_in_row(row)
                region_group = self.get_value(row, "区域")

                target_version, auto_generated = self.build_target_version(source_version, full_target_version, version_change)
                if auto_generated:
                    auto_target_version_count += 1

                ota_type = self.map_ota_type(ota_type_raw)
                upgrade_url, zip_filename = self.build_upgrade_url(source_upgrade_url, aws_suffix)
                if not zip_filename:
                    self.log_message(f"  ⚠ 第 {index} 行未能识别 ZIP 文件名；请检查源 Excel 是否包含 整包地址/升级包地址 且值以 .zip 结尾。")

                output_rows.append({
                    "设备类型代码": device_code,
                    "设备类型名称": device_name,
                    "特征码": self.validate_feature_code(feature_code),
                    "内部机型信息": self.get_value(row, "机型信息"),
                    "设备扩展信息": self.get_value(row, "机器扩展码"),
                    "OTA类型": ota_type,
                    "源版本": source_version,
                    "目标完整版本": full_target_version,
                    "目标版本": target_version,
                    "品牌组": self.get_value(row, "品牌"),
                    "区域组": region_group,
                    "MAC组": effective_mac_group,
                    "定向组": "",
                    "升级文件地址": upgrade_url,
                    "升级文件的MD5值": self.get_value(row, "升级文件的MD5值"),
                    "文件大小": self.clean_file_size(self.get_value(row, "文件大小")),
                    "SHA256": self.get_value(row, "SHA256"),
                    "EULA文件地址": self.get_value(row, "EULA文件地址"),
                })

                self.log_message("")
                self.log_message(f"第 {index} 行")
                self.log_message(f"  源 Excel 行号：{row.get('_row_number', '')}")
                self.log_message(f"  内部机型信息：{self.get_value(row, '机型信息')}")
                self.log_message(f"  设备扩展信息：{self.get_value(row, '机器扩展码')}")
                self.log_message(f"  OTA类型：{ota_type_raw} -> {ota_type}")
                self.log_message(f"  源版本：{source_version}")
                self.log_message(f"  目标完整版本：{full_target_version}")
                self.log_message(f"  目标版本：{target_version} {'(自动生成)' if auto_generated else '(来自版本变化列)'}")
                self.log_message(f"  品牌组：{self.get_value(row, '品牌')}")
                self.log_message(f"  区域组：{region_group}")
                self.log_message(f"  MAC组：{effective_mac_group}")
                self.log_message(f"  Source 映射升级包地址：{mapped_upgrade_url}")
                self.log_message(f"  Source 实际 ZIP 地址/路径：{source_upgrade_url}")
                self.log_message(f"  ZIP 文件：{zip_filename}")
                self.log_message(f"  升级文件地址：{upgrade_url}")
                self.log_message(f"  MD5：{self.get_value(row, '升级文件的MD5值')}")
                self.log_message(f"  文件大小：{self.clean_file_size(self.get_value(row, '文件大小'))}")
                self.log_message(f"  SHA256：{self.get_value(row, 'SHA256')}")
                self.validate_output_row(index, output_rows[-1])

            output_path = self.generate_output_path(source_file_path, device_code, len(output_rows))
            wb = self.create_workbook(output_rows, ota_messages=ota_messages)
            wb.save(output_path)
            self.last_generated_output_path = output_path

            self.log_message("")
            self.log_message("========== 生成结果 ==========")
            self.log_message(f"生成表格行数：{len(output_rows)}")
            self.log_message(f"自动生成目标版本行数：{auto_target_version_count}")
            self.log_message(f"升级提示语写入行数：{len(ota_messages)}")
            self.log_message(f"区域组复制状态：已从 Source Excel 的“区域”字段复制到输出“区域组”（如源文件未识别该列则为空）。")
            self.log_message(f"输出文件路径：{output_path}")
            self.log_message("模版文件生成成功。")

            if start_zip_after:
                self.status_var.set("模版完成，开始下载/上传ZIP...")
                self.log_message("")
                self.log_message("自动执行：开始后台下载 ZIP 文件；全部下载成功后将上传到 AWS。")
                worker = threading.Thread(
                    target=self.download_all_ftp_files_worker,
                    args=(source_file_path, device_code, True, pasted_text),
                    daemon=True
                )
                worker.start()
            else:
                self.stop_animation("模版生成成功")
                self.log_message(f"生成完成，Excel 文件已输出到：{output_path}")
                self.show_generated_file_prompt(output_path)

        except Exception as e:
            self.stop_animation("生成失败")
            self.log_message("")
            self.log_message("生成失败。")
            self.log_message(str(e))
            self.log_message(traceback.format_exc())
            self.show_center_error("错误", str(e))


# ==========================================================
# Main
# ==========================================================

def main():
    root = tk.Tk()
    OTATemplateApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
