"""Headless engine: reuse v3's proven business logic without the Tkinter UI.

The v3 `OTATemplateApp` mixes pure logic with GUI code in one class. Rather than
risk re-typing ~3800 lines, we instantiate it *headless* (no `__init__`, no
widgets) and call only its pure methods, redirecting its file-path accessors and
its logger to web-friendly sinks via `config`.

This gives the web app byte-for-byte identical parsing/standardization/Rule ID
behaviour as v3, and lets v3's whole regression suite keep guarding it.
"""
from __future__ import annotations

import os
import re
from datetime import datetime
from typing import Callable

from openpyxl import load_workbook

from . import config
from .legacy_engine import OTATemplateApp, AWS_PREFIX


class OTAEngine:
    """A thin, headless facade over the ported v3 logic."""

    def __init__(self, logger: Callable[[str], None] | None = None):
        config.ensure_dirs()
        self._logs: list[str] = []
        self._external_logger = logger

        app = OTATemplateApp.__new__(OTATemplateApp)  # no __init__, no widgets

        # Redirect the legacy loggers to our capture buffer.
        app.log_message = self._capture_log  # type: ignore[method-assign]
        app.thread_log = self._capture_log   # type: ignore[method-assign]

        # Redirect resource / template / history paths to v4 config.
        app.user_ota_message_templates_file = config.MESSAGE_TEMPLATES_FILE
        app.get_resource_path = self._resource_path  # type: ignore[method-assign]

        self.app = app

    # -- logging ----------------------------------------------------------
    def _capture_log(self, message: str) -> None:
        line = str(message)
        self._logs.append(line)
        if self._external_logger:
            self._external_logger(line)

    @property
    def logs(self) -> list[str]:
        return list(self._logs)

    def _resource_path(self, filename: str) -> str:
        """Resolve bundled data files (templates/history seed) from data dir."""
        candidate = os.path.join(str(config.DATA_DIR), filename)
        if os.path.exists(candidate):
            return candidate
        # Fall back to the legacy file next to the module (if present).
        return os.path.join(os.path.dirname(__file__), filename)

    # -- parsing ----------------------------------------------------------
    # Column headers that mark a vertical (one-field-per-line) OTA table.
    _VERTICAL_HEADERS = [
        "机型", "扩展信息", "源版本", "源版本地址", "目标版本", "版本变化",
        "整包地址", "差分包地址", "升级文件MD5值", "升级文件md5值",
        "升级包大小(byte)", "升级包大小", "MAC", "场测MAC", "测试MAC",
        "品牌", "OTA类型", "区域", "最大检测数", "差分包",
    ]

    def _is_vertical_header_line(self, line: str) -> bool:
        norm = self.app.normalize_header(line)
        return any(self.app.normalize_header(h) == norm for h in self._VERTICAL_HEADERS)

    def reshape_vertical_paste(self, text: str) -> str:
        """Convert a single-column (one-field-per-line) paste into tab rows.

        Excel content copied from a single selected column arrives with every
        field — header names, each cell value, and label/value metadata — on its
        own line. The normal parser expects tab-delimited rows, so this detects
        the vertical shape and rebuilds proper rows before parsing. Returns the
        text unchanged when it is not a vertical layout.
        """
        raw_lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
        lines = [ln.rstrip() for ln in raw_lines]

        # 0) Some pastes are DOUBLE-SPACED (every field followed by a blank line).
        #    In the OTA-table region that spacing is noise — collapse it so field
        #    detection works. We only touch lines up to the language section
        #    (first tab-delimited "语种\t描述" style line), leaving blank lines
        #    INSIDE quoted multi-line descriptions untouched.
        tab_start = next((i for i, ln in enumerate(lines) if "\t" in ln), len(lines))
        head = [ln for ln in lines[:tab_start] if ln.strip() != ""]
        lines = head + lines[tab_start:]

        # 1) Locate the header block: a run of short, non-value lines of which
        #    enough are known OTA headers. Tolerate unknown header names (e.g.
        #    源版本地址 / 场测MAC) that sit between known ones, so the run does not
        #    break on columns we don't enumerate.
        def is_data_value(s: str) -> bool:
            """A line that is clearly a data cell, never a header name."""
            s = s.strip()
            if not s:
                return False
            return bool(
                self.app.looks_like_model(s)
                or self.app.looks_like_version(s)
                or self.app.looks_like_version_change(s)
                or self.app.looks_like_md5(s)
                or self.app.looks_like_size(s)
                or self.app.looks_like_feature_code(s)
                or "://" in s
            )

        def is_header_ish(s: str) -> bool:
            """Short label-like line that is not a data value."""
            s = s.strip()
            return bool(s) and len(s) <= 12 and not is_data_value(s)

        # The header block starts at the first known OTA header line and runs
        # through the following header-ish lines, STOPPING at the first clear
        # data value (e.g. the first 机型 code). This keeps columns like
        # 源版本地址 / 场测MAC that we don't enumerate, without swallowing data.
        header_start = header_end = -1
        i = 0
        while i < len(lines):
            if lines[i].strip() and self._is_vertical_header_line(lines[i]):
                start = i
                end = i
                known = 1
                j = i + 1
                while j < len(lines):
                    lj = lines[j].strip()
                    if not lj or is_data_value(lj):
                        break
                    if self._is_vertical_header_line(lines[j]):
                        end = j; known += 1
                    elif is_header_ish(lj):
                        end = j
                    else:
                        break
                    j += 1
                if known >= 6:
                    header_start, header_end = start, end
                    break
                i = j + 1
            else:
                i += 1
        if header_start < 0:
            return text  # not a vertical OTA table

        headers = [lines[k].strip() for k in range(header_start, header_end + 1)]
        n = len(headers)
        mac_index = next(
            (idx for idx, h in enumerate(headers)
             if self.app.normalize_header(h) in ("mac", "场测mac", "测试mac")), None
        )

        # 2) Find where the OTA table ends (metadata/message tail starts).
        tail_start = len(lines)
        for k in range(header_end + 1, len(lines)):
            norm = self.app.normalize_header(lines[k])
            if norm in ("设备类型代码(方案号)", "设备类型代码", "featurecode") or "语种" in norm or "language" in norm:
                tail_start = k
                break

        # 3) Split the data region into rules by anchoring on a new 机型 value.
        #    Each rule's fields are collected (blanks preserved as empty cells),
        #    then padded/trimmed to the header width. A MAC column that arrived as
        #    two hex lines is merged.
        data_region = [lines[k] for k in range(header_end + 1, tail_start)]
        # Drop leading blank lines.
        while data_region and not data_region[0].strip():
            data_region.pop(0)

        def is_new_model(s: str) -> bool:
            s = s.strip()
            if not s or not self.app.looks_like_model(s):
                return False
            # Versions (e.g. 0100AP0920) also match the loose model pattern; a new
            # rule anchor must NOT be a version / version-change / url / hash.
            if (self.app.looks_like_version(s) or self.app.looks_like_version_change(s)
                    or self.app.looks_like_md5(s) or "://" in s):
                return False
            return True

        rules: list[list[str]] = []
        current: list[str] = []
        for ln in data_region:
            s = ln.strip()
            if is_new_model(s) and current:
                rules.append(current)
                current = []
            # Merge consecutive MAC-only lines into one field: a bare hex group
            # (e.g. 3400ffb789 then 3400ffb790) copied on separate lines belongs
            # to a single MAC cell. Only merge true MAC hex (contains a-f letters)
            # so pure-decimal values like a file size (230066209) are never
            # absorbed into MAC.
            def is_mac_hex(v: str) -> bool:
                v = v.strip().rstrip("、,，")
                return (bool(re.fullmatch(r"(?i)[a-f0-9]{8,14}", v))
                        and bool(re.search(r"(?i)[a-f]", v))
                        and not self.app.looks_like_size(v))
            prev = current[-1] if current else ""
            # Merge a MAC continuation: either the previous MAC field ends with a
            # separator (c...、) and this line is another hex group, or both lines
            # are bare MAC hex groups split across lines.
            if current and s and is_mac_hex(s) and (
                prev.rstrip().endswith(("、", ",", "，")) or is_mac_hex(prev)
            ):
                sep = "" if prev.rstrip().endswith(("、", ",", "，")) else "、"
                current[-1] = prev + sep + s
                continue
            current.append(s)
        if current:
            rules.append(current)

        data_rows = []
        for fields in rules:
            if len(fields) < n:
                fields = fields + [""] * (n - len(fields))
            elif len(fields) > n:
                fields = fields[:n]
            data_rows.append(fields)

        if not data_rows:
            return text

        # 3) Rebuild: reshaped HEAD metadata (before the table) + header row +
        #    data rows + reshaped TAIL (metadata pairs + language table), so
        #    device detection and message parsing both work no matter where the
        #    设备类型代码 / featurecode block sits.
        out = self._reshape_vertical_tail(lines[:header_start])
        if out:
            out.append("")
        out.append("\t".join(headers))
        out += ["\t".join(r) for r in data_rows]
        out.append("")
        out += self._reshape_vertical_tail(lines[tail_start:])
        return "\n".join(out)

    def _reshape_vertical_tail(self, tail_lines: list[str]) -> list[str]:
        """Reshape the metadata + language section of a vertical paste.

        Handles three language shapes, all producing a clean '语种\t描述' table:
        - 语种\t"multi-line quoted description" (quote may span blank lines)
        - 语种 label on its own line, description on the following lines
        - a 语种/Language header followed by either of the above
        Also pairs metadata labels (设备类型代码（方案号）/featurecode) with their
        value on the next line.
        """
        out: list[str] = []
        in_messages = False
        i = 0
        n = len(tail_lines)
        meta_labels = ("设备类型代码(方案号)", "设备类型代码", "featurecode")

        def emit_lang_header_once():
            if not any(x == "语种\t描述" for x in out):
                out.append("语种\t描述")

        while i < n:
            ln = tail_lines[i].strip()
            if not ln:
                out.append("")
                i += 1
                continue
            norm = self.app.normalize_header(ln)

            if "语种" in norm or "language" in norm:
                emit_lang_header_once()
                in_messages = True
                i += 1
                continue

            # A '语种\t"..."' line (tab-delimited label + quoted description).
            if "\t" in tail_lines[i]:
                label, _, rest = tail_lines[i].partition("\t")
                if self.app.normalize_message_language(label.strip()):
                    in_messages = True
                    emit_lang_header_once()
                    body = rest.strip()
                    # Absorb a multi-line quoted description across blank lines
                    # until the closing quote.
                    if body.startswith('"') and not (body.endswith('"') and len(body) > 1):
                        parts = [body.lstrip('"')]
                        j = i + 1
                        while j < n:
                            seg = tail_lines[j]
                            if seg.strip().endswith('"'):
                                parts.append(seg.strip().rstrip('"'))
                                j += 1
                                break
                            if seg.strip():
                                parts.append(seg.strip())
                            j += 1
                        desc = "\n".join(p for p in parts if p != "").strip()
                        i = j
                    else:
                        desc = body.strip('"').strip()
                        i += 1
                    first = desc.split("\n")[0] if desc else ""
                    out.append(f"{label.strip()}\t{first}")
                    for extra in desc.split("\n")[1:]:
                        out.append(f"\t{extra}")
                    continue

            if not in_messages:
                if norm in meta_labels:
                    j = i + 1
                    while j < n and not tail_lines[j].strip():
                        j += 1
                    value = tail_lines[j].strip() if j < n else ""
                    out.append(f"{ln}\t{value}")
                    i = j + 1
                    continue
                out.append(ln)
                i += 1
                continue

            # Plain vertical language block: label line then description lines.
            if self.app.normalize_message_language(ln):
                emit_lang_header_once()
                desc: list[str] = []
                j = i + 1
                while j < n:
                    nxt = tail_lines[j].strip()
                    if not nxt:
                        j += 1
                        continue
                    if self.app.normalize_message_language(nxt):
                        break
                    desc.append(nxt)
                    j += 1
                first = desc[0] if desc else ""
                out.append(f"{ln}\t{first}")               # 语种 \t 第一行描述
                for extra in desc[1:]:
                    out.append(f"\t{extra}")               # 续行：空语种列 + 描述行
                i = j
                continue
            i += 1
        return out

    def parse_pasted_text(self, text: str) -> dict:
        """Parse pasted email/table content into the JSON intermediate format."""
        reshaped = self.reshape_vertical_paste(text)
        return self.app.build_source_json_from_text(reshaped, use_gui_log=False)

    def parse_excel(self, file_path: str) -> dict:
        return self.app.build_source_json_from_excel(file_path, use_gui_log=False)

    # -- validation -------------------------------------------------------
    def _validate_generation_inputs(self, *, rows, device_code, feature_code, mac_group):
        """Raise ValueError(多行) if required fields are missing/malformed.

        Collects ALL problems so the user can fix them in one pass. Format rules
        are lenient where device naming legitimately varies (e.g. AWOL models),
        strict where a fixed format is required (feature code = 18+ digits).
        """
        import re as _re
        app = self.app
        errors: list[str] = []

        code = str(device_code or "").strip()
        if not code:
            errors.append("设备类型代码不能为空。")
        elif _re.search(r"\s", code):
            errors.append(f"设备类型代码不能包含空格：{code!r}。")
        elif not _re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_\-().]*", code):
            errors.append(f"设备类型代码格式不正确（只能是字母/数字及 _-().）：{code!r}。")

        fc = str(feature_code or "").strip()
        if fc and not app.looks_like_feature_code(fc):
            errors.append(f"Feature Code / 特征码格式不正确（应为至少 18 位数字）：{fc!r}。")

        mac = str(mac_group or "").strip()
        if not mac:
            errors.append("MAC 地址组不能为空（可填写，或使用“测试+设备类型代码”）。")

        if not rows:
            errors.append("未解析到任何 OTA 数据行，无法生成模版。")

        # Per-row identity checks: model and extension must be present and the
        # extension must look like an extension code, not an empty/garbage value.
        for idx, row in enumerate(rows, start=1):
            model = str(app.get_value(row, "机型信息") or "").strip()
            ext = str(app.get_value(row, "机器扩展码") or "").strip()
            if not model:
                errors.append(f"第 {idx} 行：机型信息为空。")
            if not ext:
                errors.append(f"第 {idx} 行：设备扩展码为空。")
            elif model and ext == model:
                errors.append(f"第 {idx} 行：设备扩展码与机型信息相同，疑似解析错位：{ext!r}。")
            elif _re.fullmatch(r"\d+", ext):
                errors.append(f"第 {idx} 行：设备扩展码格式不正确（不应是纯数字）：{ext!r}。")

        if errors:
            raise ValueError("生成模版前检测到以下问题，请修正后重试：\n- " + "\n- ".join(errors))

    # -- generation -------------------------------------------------------
    def generate_excel(
        self,
        source_json: dict,
        *,
        device_code: str = "",
        device_name: str = "",
        feature_code: str = "",
        mac_group: str = "",
        output_dir: str | None = None,
    ) -> dict:
        """Build the OTA import workbook from a parsed source_json (headless).

        Mirrors v3's generate_excel_template row-building, minus all UI vars and
        confirmation dialogs. Returns the output path plus a summary dict.
        """
        app = self.app
        rows = source_json.get("rows", [])
        mapping = source_json.get("mapping", {})
        source_ota_messages = source_json.get("messages", [])

        aws_suffix = app.get_default_aws_suffix(device_code)
        ota_messages, ota_message_source = app.select_ota_messages_for_generation(
            source_ota_messages,
            device_code,
            source_type=source_json.get("source_type", ""),
            prompt_user=False,  # never prompt in a web/server context
        )

        source_mac_group = next(
            (app.get_value(row, "MAC") for row in rows if app.get_value(row, "MAC")), ""
        )
        effective_mac_group = mac_group or source_mac_group

        # Validate inputs before building the workbook; raise a combined, readable
        # error so the user can fix everything at once.
        self._validate_generation_inputs(
            rows=rows,
            device_code=device_code,
            feature_code=feature_code,
            mac_group=effective_mac_group,
        )

        output_rows = []
        auto_target_version_count = 0
        for row in rows:
            source_version = app.get_value(row, "源版本")
            full_target_version = app.get_value(row, "目标版本")
            version_change = app.get_value(row, "版本变化")
            ota_type_raw = app.get_value(row, "OTA类型")
            source_upgrade_url = app.find_zip_source_url_in_row(row)
            region_group = app.get_value(row, "区域")

            target_version, auto_generated = app.build_target_version(
                source_version, full_target_version, version_change
            )
            if auto_generated:
                auto_target_version_count += 1
            ota_type = app.map_ota_type(ota_type_raw)
            upgrade_url, _zip = app.build_upgrade_url(source_upgrade_url, aws_suffix)

            output_rows.append({
                "设备类型代码": device_code,
                "设备类型名称": device_name,
                "特征码": app.validate_feature_code(feature_code),
                "内部机型信息": app.get_value(row, "机型信息"),
                "设备扩展信息": app.get_value(row, "机器扩展码"),
                "OTA类型": ota_type,
                "源版本": source_version,
                "目标完整版本": full_target_version,
                "目标版本": target_version,
                "品牌组": app.get_value(row, "品牌"),
                "区域组": region_group,
                "MAC组": effective_mac_group,
                "定向组": "",
                "升级文件地址": upgrade_url,
                "升级文件的MD5值": app.get_value(row, "升级文件的MD5值"),
                "文件大小": app.clean_file_size(app.get_value(row, "文件大小")),
                "SHA256": app.get_value(row, "SHA256"),
                "EULA文件地址": app.get_value(row, "EULA文件地址"),
            })

        out_dir = output_dir or str(config.DATA_DIR / "outputs")
        os.makedirs(out_dir, exist_ok=True)
        safe_device = app.safe_filename_part(device_code or "OTA")
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(out_dir, f"{safe_device}_{len(output_rows)}_{stamp}.xlsx")

        wb = app.create_workbook(output_rows, ota_messages=ota_messages)
        wb.save(output_path)

        return {
            "output_path": output_path,
            "row_count": len(output_rows),
            "auto_target_version_count": auto_target_version_count,
            "message_count": len(ota_messages),
            "message_source": ota_message_source,
            "mapping": mapping,
            "mac_group": effective_mac_group,
        }

    # -- device presets ---------------------------------------------------
    def list_device_presets(self) -> list[dict]:
        """Return the known device-code presets (from the seed/history)."""
        seed = self.app.load_device_history_seed()
        return [
            {
                "device_code": r.get("device_code", ""),
                "device_name": r.get("device_name", ""),
                "feature_code": r.get("feature_code", ""),
                "mac_group": r.get("mac_group", ""),
            }
            for r in seed
            if r.get("device_code")
        ]

    def resolve_device_info_from_text(self, text: str) -> dict:
        """Robustly derive device info from pasted text.

        Priority for each field:
          1. Value explicitly present in the pasted text (labels or 方案号 rows).
          2. Seed preset for the detected device code (fills name / feature code).
        The suggested MAC group is always '测试' + device_code unless the pasted
        text explicitly provided a MAC group.

        Always returns a dict with the four keys (possibly empty) plus
        `detected` listing which fields came from the pasted text.
        """
        text = str(text or "")
        # Normalize a vertical (one-field-per-line) paste first so label/value
        # metadata like "设备类型代码（方案号）\nMT9900NAAnd4KUHD" is seen as a pair.
        text = self.reshape_vertical_paste(text) if text.strip() else text
        pasted = self.app.extract_form_fields_from_text(text) if text.strip() else {}
        seed = self.app.load_device_history_seed()

        device_code = str(pasted.get("device_code", "") or "").strip()
        pasted_feature = str(pasted.get("feature_code", "") or "").strip()

        # Resolve the preset by device code; if the code is missing but a Feature
        # Code was found, reverse-look-up the device by its (unique) feature code.
        preset = {}
        if device_code:
            preset = next((r for r in seed if r.get("device_code") == device_code), {})
        if not preset and pasted_feature:
            preset = next((r for r in seed if str(r.get("feature_code", "")).strip() == pasted_feature), {})
            if preset and not device_code:
                device_code = str(preset.get("device_code", "") or "").strip()

        device_name = str(pasted.get("device_name", "") or "").strip() or str(preset.get("device_name", "") or "").strip()
        feature_code = pasted_feature or str(preset.get("feature_code", "") or "").strip()

        pasted_mac = str(pasted.get("mac_group", "") or "").strip()
        if pasted_mac:
            mac_group = pasted_mac
        elif device_code:
            # Suggested name is always 测试 + device type code.
            mac_group = f"测试{device_code}"
        else:
            mac_group = ""

        return {
            "device_code": device_code,
            "device_name": device_name,
            "feature_code": feature_code,
            "mac_group": mac_group,
            "detected": {k: bool(v) for k, v in pasted.items()},
        }

    def device_preset(self, device_code: str) -> dict:
        """Look up a device code and return its fixed fields.

        device_name / feature_code come from the preset (may be empty). The MAC
        group defaults to '测试' + device_code when the preset has none.
        """
        code = str(device_code or "").strip()
        record = next(
            (r for r in self.app.load_device_history_seed() if r.get("device_code") == code),
            {},
        )
        mac_group = str(record.get("mac_group", "") or "").strip()
        if not mac_group and code:
            mac_group = f"测试{code}"
        return {
            "device_code": code,
            "device_name": str(record.get("device_name", "") or "").strip(),
            "feature_code": str(record.get("feature_code", "") or "").strip(),
            "mac_group": mac_group,
        }

    # -- transfer helpers -------------------------------------------------
    def extract_ftp_zip_urls(self, source_json: dict) -> list[str]:
        """Collect unique FTP .zip URLs from a parsed source_json's rows."""
        urls: list[str] = []
        for row in source_json.get("rows", []):
            url = self.app.find_zip_source_url_in_row(row)
            if url and str(url).lower().startswith("ftp://") and ".zip" in str(url).lower():
                urls.append(str(url).strip())
        return list(dict.fromkeys(urls))

    def extract_ftp_zip_urls_from_excel(self, file_path: str) -> list[str]:
        """Scan a workbook for FTP .zip URLs (header-independent)."""
        urls = self.app.scan_ftp_zip_urls_from_excel(file_path)
        return [u for u in urls if str(u).lower().startswith("ftp://")]

    def preview_excel(self, file_path: str, max_rows: int = 200) -> dict:
        """Read a generated workbook into a header + rows structure for preview.

        Reads the first worksheet (the OTA 升级规则 sheet). Returns the header row
        and up to `max_rows` data rows so the UI can render the updated file
        in-place instead of forcing a re-download.
        """
        return self.app.read_excel_preview(file_path, max_rows=max_rows)

    # -- Rule ID ----------------------------------------------------------
    def parse_rule_id_range(self, raw_value: str) -> list[int]:
        return self.app.parse_rule_id_range(raw_value)

    def update_rule_id_columns(self, file_path: str, rule_id_range: str) -> dict:
        """Validate the input, insert the 3 columns, return the summary.

        If the user enters a SINGLE Rule ID (no range), auto-increment it by +1
        per data row to fill the whole table. An explicit range must match the
        row count exactly.
        """
        rule_ids = self.app.parse_rule_id_range(rule_id_range)
        auto_increment = len(rule_ids) == 1  # single start id -> auto-fill
        row_count, assigned_ids, records = self.app.insert_rule_id_columns(
            file_path, rule_ids, auto_increment=auto_increment
        )
        summary = self.app.build_rule_id_deployment_summary(records, file_path=file_path)
        return {
            "row_count": row_count,
            "assigned_ids": assigned_ids,
            "records": records,
            "summary": summary,
        }
