"""场测 MAC 与 CMS「维度管理 → MAC组管理」的 MAC 组（纯逻辑部分）。

CMS 的 MAC 组只存 MAC 的『后 8 位』（十六进制、小写），每段是一个区间『起-止』，例如
    7de6d0e5-7de6d0f4,7d71d971-7d71d971
单个地址就是起止相同的区间。

转换：cc:12:28:cd:db:d8 → 去掉分隔符 cc1228cddbd8 → 取后 8 位 28cddbd8。
"""
from __future__ import annotations

import re

# 可识别的 MAC 写法：
#   完整 MAC：6 组两位十六进制，分隔符可以是 : - . 空格 或没有分隔符（12 位连写）
#   只给后 8 位：8 位十六进制连写（'b7fc8ce6'），或 4 组两位用 : / - 分隔（'b7:fc:8c:e6'）
# 前后不能紧挨着十六进制位，也不能是『分隔符 + 十六进制位』（避免从更长的串里截一段，比如 5 组的残缺 MAC）。
_MAC_RE = re.compile(
    r"(?i)(?<![0-9a-f])(?<![0-9a-f][:\-])"
    r"((?:[0-9a-f]{2}[:\-. ]){5}[0-9a-f]{2}|[0-9a-f]{12}|(?:[0-9a-f]{2}[:\-]){3}[0-9a-f]{2}|[0-9a-f]{8})"
    r"(?![0-9a-f])(?![:\-][0-9a-f])")


def mac_tail8(mac: str) -> str:
    """一个 MAC → 后 8 位（小写）。'cc:12:28:cd:db:d8' → '28cddbd8'。

    也接受已经是 8 位的写法（'28CDDBD8' → '28cddbd8'）。看不懂返回空串。
    """
    h = re.sub(r"(?i)[^0-9a-f]", "", str(mac or ""))
    if len(h) == 12 or len(h) == 8:
        return h[-8:].lower()
    return ""


def extract_macs(cell: str) -> list[str]:
    """从一个『场测MAC』单元格里取出所有 MAC，返回后 8 位列表（去重保序）。

    单元格可能是：'cc:12:28:cd:db:d8' / '1）cc:12:28:db:91:ec\\n2）a0:00:4c:c8:36:d9' /
    直接给后 8 位 'b7fc8ce6、b7fc8ce7' / 逗号、顿号、空格分隔 / '空着为默认不设置'（没有 MAC，返回 []）。
    """
    out: list[str] = []
    for m in _MAC_RE.finditer(str(cell or "")):
        t = mac_tail8(m.group(1))
        if t and t not in out:
            out.append(t)
    return out


def parse_ranges(text: str) -> list[tuple[str, str]]:
    """CMS 列表『MAC组』列 → [(起, 止), ...]（后 8 位、小写）。'7de6d0e5-7de6d0f4,63553bfe-63553bfe'。"""
    out = []
    for part in re.split(r"[,，\s]+", str(text or "").strip()):
        if not part:
            continue
        a, _, b = part.partition("-")
        a, b = a.strip().lower(), (b or a).strip().lower()
        if re.fullmatch(r"[0-9a-f]{8}", a) and re.fullmatch(r"[0-9a-f]{8}", b):
            out.append((a, b))
    return out


def in_ranges(tail8: str, ranges: list[tuple[str, str]]) -> bool:
    """后 8 位是否落在任一区间内（按十六进制数值比较，含两端）。"""
    v = int(tail8, 16)
    return any(int(a, 16) <= v <= int(b, 16) for a, b in ranges)


def missing_macs(wanted: list[str], ranges: list[tuple[str, str]]) -> list[str]:
    """要求的 MAC（后 8 位）里，不在组内的那些。"""
    return [m for m in wanted if not in_ranges(m, ranges)]


def group_name(device_code: str) -> str:
    """测试 MAC 组的名称：『测试』+ 设备类型代码。"""
    return "测试" + str(device_code or "").strip()


def macs_from_rows(rows: list[dict]) -> dict[str, list[str]]:
    """解析后的数据行 → {后8位: [出现在哪些行的机型]}。"""
    out: dict[str, list[str]] = {}
    for r in rows:
        for m in extract_macs(r.get("MAC", "")):
            out.setdefault(m, []).append(str(r.get("机型信息", "") or "").strip())
    return out


# 去掉合法 MAC 之后，剩下的『像十六进制串』（含 : - . 分隔）有 ≥6 个十六进制位 → 看不懂的 MAC 写法
_HEXISH_RE = re.compile(r"(?i)[0-9a-f](?:[0-9a-f:\-.]*[0-9a-f])?")
# 原文全文扫描只认带 : 或 - 分隔的标准写法（12 位连写在 URL / 哈希里容易误判）
_SEP_MAC_RE = re.compile(r"(?i)(?<![0-9a-f:\-])((?:[0-9a-f]{2}:){5}[0-9a-f]{2}|(?:[0-9a-f]{2}-){5}[0-9a-f]{2})(?![0-9a-f:\-])")


def unrecognized_parts(cell: str) -> list[str]:
    """单元格里『像 MAC 但不是完整 12 位 MAC』的片段（如只写了 8 位、少一位）。说明文字（无十六进制串）返回 []。"""
    rest = _MAC_RE.sub(" ", str(cell or ""))
    out = []
    for m in _HEXISH_RE.finditer(rest):
        tok = m.group(0)
        if len(re.sub(r"(?i)[^0-9a-f]", "", tok)) >= 6:
            out.append(tok)
    return out


def scan_rows(rows: list[dict]) -> list[dict]:
    """逐行列出场测 MAC：[{行号(从1起), 机型, 原文, macs(后8位), bad(看不懂的片段)}]。"""
    out = []
    for i, r in enumerate(rows, 1):
        raw = str(r.get("MAC", "") or "").strip()
        out.append({"row": i, "model": str(r.get("机型信息", "") or "").strip(), "raw": raw,
                    "macs": extract_macs(raw), "bad": unrecognized_parts(raw)})
    return out


def scan_text_macs(text: str) -> list[str]:
    """在整份原文里找标准写法的 MAC（xx:xx:xx:xx:xx:xx / xx-xx-…），返回后 8 位（去重保序）。"""
    out: list[str] = []
    for m in _SEP_MAC_RE.finditer(str(text or "")):
        t = mac_tail8(m.group(1))
        if t and t not in out:
            out.append(t)
    return out



def _selftest() -> int:
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    check("cc:12:28:cd:db:d8 → 28cddbd8", mac_tail8("cc:12:28:cd:db:d8") == "28cddbd8")
    check("大写 / 横线 / 连写都能转", mac_tail8("CC-12-28-CD-DB-D8") == "28cddbd8"
          and mac_tail8("CC1228CDDBD8") == "28cddbd8" and mac_tail8("28CDDBD8") == "28cddbd8")
    check("看不懂的返回空", mac_tail8("空着为默认不设置") == "" and mac_tail8("12:34") == "")
    check("单元格里多个编号 MAC", extract_macs("1）cc:12:28:db:91:ec\n2）a0:00:4c:c8:36:d9") == ["28db91ec", "4cc836d9"])
    check("逗号 / 顿号 / 重复", extract_macs("cc:12:28:cd:db:d8、cc:12:28:cd:db:d8,a0:00:4c:c8:36:d9")
          == ["28cddbd8", "4cc836d9"])
    check("没有 MAC 的单元格", extract_macs("空着为默认不设置") == [] and extract_macs("") == [])
    check("不把 MD5 / feature code 当成 MAC", extract_macs("debfa7e56660272595ba7e257eb3ce44") == []
          and extract_macs("861003009000010000000101") == [])
    rg = parse_ranges("7de6d0e5-7de6d0f4,7d71d971-7d71d971,28cddbd8-28cddbd8")
    check("区间解析", rg == [("7de6d0e5", "7de6d0f4"), ("7d71d971", "7d71d971"), ("28cddbd8", "28cddbd8")])
    check("区间内 / 端点 / 单个地址", in_ranges("7de6d0ea", rg) and in_ranges("7de6d0f4", rg) and in_ranges("28cddbd8", rg))
    check("区间外", not in_ranges("7de6d0f5", rg) and missing_macs(["28cddbd8", "4cc836d9"], rg) == ["4cc836d9"])
    check("组名", group_name("A311D2AndrWE6GE") == "测试A311D2AndrWE6GE")
    check("从数据行收集", macs_from_rows([{"MAC": "", "机型信息": "a"}, {"MAC": "cc:12:28:cd:db:d8", "机型信息": "65WE6GE(01)"}])
          == {"28cddbd8": ["65WE6GE(01)"]})
    check("直接给后 8 位（顿号 / 逗号 / 换行 / 大写）", extract_macs("b7fc8ce6、b7fc8ce7") == ["b7fc8ce6", "b7fc8ce7"]
          and extract_macs("B7FC8CE6,\nb7fc8ce7") == ["b7fc8ce6", "b7fc8ce7"]
          and unrecognized_parts("b7fc8ce6、b7fc8ce7") == [])
    check("后 8 位用冒号 / 横线分隔", extract_macs("b7:fc:8c:e6") == ["b7fc8ce6"]
          and extract_macs("1）B7-FC-8C-E6") == ["b7fc8ce6"])
    check("完整 MAC 与后 8 位混写", extract_macs("cc:12:28:cd:db:d8、b7fc8ce7") == ["28cddbd8", "b7fc8ce7"])
    check("看不懂：少一位 / 7 位 / 10 位 / 5 组残缺",
          unrecognized_parts("cc:12:28:cd:db:d") == ["cc:12:28:cd:db:d"]
          and unrecognized_parts("b7fc8ce") == ["b7fc8ce"] and unrecognized_parts("12b7fc8ce6") == ["12b7fc8ce6"]
          and extract_macs("cc:12:28:cd:db") == [] and unrecognized_parts("cc:12:28:cd:db") == ["cc:12:28:cd:db"])
    check("后 8 位区间写法不猜，报看不懂", extract_macs("b7fc8ce6-b7fc8ce7") == []
          and unrecognized_parts("b7fc8ce6-b7fc8ce7") == ["b7fc8ce6-b7fc8ce7"])
    check("看得懂的不报", unrecognized_parts("1）cc:12:28:db:91:ec\n2）a0:00:4c:c8:36:d9") == []
          and unrecognized_parts("1.cc:12:28:db:91:ec") == [])
    check("说明文字不报", unrecognized_parts("空着为默认不设置") == [] and unrecognized_parts("无 / N/A / default") == [])
    sr = scan_rows([{"MAC": "cc:12:28:cd:db:d8", "机型信息": "65"}, {"MAC": "28cddbd", "机型信息": "75"},
                    {"MAC": "", "机型信息": "86"}])
    check("逐行扫描", [r["macs"] for r in sr] == [["28cddbd8"], [], []] and sr[1]["bad"] == ["28cddbd"]
          and sr[2]["row"] == 3 and sr[2]["model"] == "86")
    check("全文扫描只认分隔写法", scan_text_macs("场测MAC：cc:12:28:cd:db:d8，备用 A0-00-4C-C8-36-D9 "
                                       "md5 debfa7e56660272595ba7e257eb3ce44 cc1228cddbd8")
          == ["28cddbd8", "4cc836d9"])
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    import sys
    sys.exit(_selftest() if "--selftest" in sys.argv[1:] else 0)
