"""OTA 规则行的字段特征检查（纯逻辑）。prep 解析后、生成 Excel 前查一次；最终检查对生成的 Excel 再查一次。

每个字段都有很明显的特征，值放错列（列错位）时一眼就能看出来：

  字段            特征                                                        不符合时
  设备类型代码    字母开头、只有字母数字、8~24 位、含数字（MT8195Andr32P1S、A311D2AndrWE6GE）；整批只有一个  ✗
  特征码          18~24 位纯数字；整批只有一个                                       ✗
  机型            字母数字开头，可带括号后缀（32P1S(9A)、HU58A6109FUWA）；
                  不能是版本号 / 纯数字 / MD5 / 地址                                ✗（只是字符少见：⚠）
  扩展码          同上，不能是版本号 / 纯数字 / MD5 / 地址                           ✗
                  通常以机型（去掉括号）开头，但不一定                               ⚠
  源 / 目标版本   4 位数字 + 1~3 个字母 + 4 位数字（0001CQ0903、0100AWP0224）或 V1.2.3      ✗
  版本变化        X_to_Y，两头和源 / 目标版本末尾对得上                              ✗
  MD5             32 位十六进制                                                    ✗
  包大小          纯数字（可带千分位逗号），大于 0                                   ✗
  升级包地址      ://…….zip                                                        ✗
  品牌 / 区域     不是数字 / 版本号 / 地址（✗）；为空 = 不限（⚠）
  跨行            同名升级包（S3 上同一地址）MD5、大小必须一样；完全重复的行（机型+扩展码+源+目标）  ✗

行数据用逻辑键：code feature model ext src tgt change url md5 size brand region（缺的键不查）。
"""
from __future__ import annotations

import re

VERSION_RE = re.compile(r"(?i)^(?:\d{4}[A-Z]{1,3}\d{4}|V\d+(?:\.\d+)+)$")
CODE_RE = re.compile(r"^(?=[A-Za-z0-9]*\d)[A-Za-z][A-Za-z0-9]{7,23}$")   # MT8195Andr32P1S、A311D2AndrWE6GE
FEATURE_RE = re.compile(r"^\d{18,24}$")
MODEL_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._\-]*(?:\([A-Za-z0-9._\-]+\))?$")
MD5_RE = re.compile(r"(?i)^[0-9a-f]{32}$")

# 引擎解析出的行（源字段名）→ 逻辑键
SOURCE_KEYS = {"机型信息": "model", "机器扩展码": "ext", "源版本": "src", "目标版本": "tgt", "版本变化": "change",
               "升级包地址": "url", "升级文件的MD5值": "md5", "文件大小": "size", "品牌": "brand", "区域": "region"}
LABEL = {"code": "设备类型代码", "feature": "特征码", "model": "机型", "ext": "扩展码", "src": "源版本",
         "tgt": "目标版本", "change": "版本变化", "url": "升级包地址", "md5": "MD5", "size": "包大小",
         "brand": "品牌", "region": "区域"}


def _s(v) -> str:
    return " ".join(str(v if v is not None else "").split())


def kind_of(v: str) -> str:
    """这个值『看起来像』什么（用来写清楚放错列时到底是什么）。"""
    v = _s(v)
    if not v:
        return "空"
    if "://" in v:
        return "地址"
    if MD5_RE.match(v):
        return "MD5"
    if VERSION_RE.match(v):
        return "版本号"
    if re.fullmatch(r"(?i)[^_\s]+_to_[^_\s]+", v):
        return "版本变化"
    if re.fullmatch(r"\d[\d,]*", v):
        return "数字"
    return ""


def _id_like(v: str) -> tuple[str, str]:
    """机型 / 扩展码：返回 (严重程度 'fail'/'warn'/'', 说明)。"""
    k = kind_of(v)
    if k:
        return "fail", (f"是{k}，列可能错位了" if k != "空" else "为空")
    if not MODEL_RE.match(v):
        return "warn", "含少见字符"
    return "", ""


def _core(model: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", re.sub(r"\([^)]*\)", "", model)).upper()


def check_rows(rows: list[dict], code: str = "", feature: str = "") -> list[tuple[str, str]]:
    """rows：逻辑键的行。返回 [(严重程度, 说明)]，严重程度 'fail' / 'warn'。code / feature 是整批的值（可空）。"""
    out: list[tuple[str, str]] = []
    if code and not CODE_RE.match(_s(code)):
        out.append(("fail", f"设备类型代码『{code}』不像设备类型代码（应如 MT8195Andr32P1S）"))
    if feature and not FEATURE_RE.match(_s(feature)):
        out.append(("fail", f"特征码『{feature}』不是 18~24 位数字"))
    for key in ("code", "feature"):
        vals = sorted({_s(r.get(key)) for r in rows if _s(r.get(key))})
        if len(vals) > 1:
            out.append(("fail", f"一批里有 {len(vals)} 个不同的{LABEL[key]}：{' / '.join(vals)}"))
        for v in vals:
            if key == "code" and not CODE_RE.match(v):
                out.append(("fail", f"设备类型代码『{v}』不像设备类型代码"))
            if key == "feature" and not FEATURE_RE.match(v):
                out.append(("fail", f"特征码『{v}』不是 18~24 位数字"))

    by_url: dict[str, set] = {}
    seen: dict[tuple, int] = {}
    for i, r in enumerate(rows, 1):
        g = lambda k: _s(r.get(k))
        where = f"第{i}行 {g('model') or '?'}"
        for key in ("model", "ext"):
            if key in r:
                sev, why = _id_like(g(key))
                if sev:
                    out.append((sev, f"{where}：{LABEL[key]}『{g(key) or '（空）'}』{why}"))
        m, e = g("model"), g("ext")
        if (m and e and not kind_of(m) and not kind_of(e) and _core(m)
                and not re.sub(r"[^A-Za-z0-9]", "", e).upper().startswith(_core(m))):
            out.append(("warn", f"{where}：扩展码『{e}』不是以机型『{m}』开头（通常是，但不一定）"))
        for key in ("src", "tgt"):
            if key in r and not VERSION_RE.match(g(key)):
                k = kind_of(g(key))
                out.append(("fail", f"{where}：{LABEL[key]}『{g(key) or '（空）'}』不像版本号"
                                    + (f"（是{k}）" if k and k != "空" else "")))
        mm = re.match(r"(?i)^(.+?)_to_(.+)$", g("change"))
        if mm and VERSION_RE.match(g("src")) and VERSION_RE.match(g("tgt")):
            a_, b_ = mm.group(1).upper(), mm.group(2).upper()
            src_ok, tgt_ok = g("src").upper().endswith(a_), g("tgt").upper().endswith(b_)
            if not (src_ok and tgt_ok):
                # 只差末尾多出的字母（如 Q0725t）：对得上前面部分 → ⚠；真对不上 → ✗
                a2, b2 = re.sub(r"[A-Z]+$", "", a_), re.sub(r"[A-Z]+$", "", b_)
                loose = g("src").upper().endswith(a2) and g("tgt").upper().endswith(b2) and a2 and b2
                out.append(("warn" if loose else "fail",
                            f"{where}：版本变化『{g('change')}』和 {g('src')} → {g('tgt')} 对不上"
                            + ("（只差末尾字母）" if loose else "")))
        if "md5" in r and not MD5_RE.match(g("md5")):
            out.append(("fail", f"{where}：MD5『{g('md5') or '（空）'}』不是 32 位十六进制"))
        if "size" in r:
            sz = g("size").replace(",", "").replace("，", "")
            if not sz.isdigit() or int(sz) <= 0:
                out.append(("fail", f"{where}：包大小『{g('size') or '（空）'}』不是正整数"))
        if "url" in r and not ("://" in g("url") and g("url").lower().endswith(".zip")):
            out.append(("fail", f"{where}：升级包地址『{g('url') or '（空）'}』不是 …/*.zip"))
        for key in ("brand", "region"):
            if key in r:
                k = kind_of(g(key))
                if k == "空":
                    out.append(("warn", f"{where}：{LABEL[key]}为空（= 不限）"))
                elif k:
                    out.append(("fail", f"{where}：{LABEL[key]}『{g(key)}』是{k}，列可能错位了"))
        if g("url"):
            # 按文件名分组：上传 S3 后路径只有『设备 / 日期 / 文件名』，同名的两个包会落到同一个地址、互相覆盖。
            name = g("url").rsplit("/", 1)[-1]
            by_url.setdefault(name, set()).add((g("md5").lower(), g("size").replace(",", "").replace("，", "")))
        if e and g("src"):
            key2 = (m.upper(), e.upper(), g("src").upper(), g("tgt").upper())
            if key2 in seen:
                out.append(("fail", f"{where}：机型 + 扩展码 {e} + 源版本 {g('src')} → {g('tgt')} 和第{seen[key2]}行完全重复"))
            else:
                seen[key2] = i
    for name, pairs in by_url.items():
        if len(pairs) > 1:
            out.append(("fail", f"两个不同的升级包同名 {name}（MD5 / 大小：" +
                        "；".join(f"{a or '?'} / {b or '?'}" for a, b in sorted(pairs)) +
                        "）：上传 S3 后会落到同一个地址、互相覆盖。请让提供方改名或分开部署"))
    return out


def from_source(rows: list[dict]) -> list[dict]:
    """引擎解析出的行 → 逻辑键的行。"""
    return [{SOURCE_KEYS[k]: v for k, v in r.items() if k in SOURCE_KEYS} for r in rows]


def _selftest() -> int:
    good = True

    def check(name, cond):
        nonlocal good
        print(f"  [{'PASS' if cond else 'FAIL'}] {name}")
        good = good and bool(cond)

    ok_row = {"model": "32P1S(9A)", "ext": "32P1S9A888", "src": "0001CQ0903", "tgt": "0001CQ1116",
              "change": "Q0903_to_Q1116", "url": "ftp://u@h/a/X.zip", "md5": "d82026de64b20398467b8c1423b19aaa",
              "size": "118,219,077", "brand": "REGZA", "region": "全球"}
    check("正常的行没有问题", check_rows([ok_row], "MT8195Andr32P1S", "861003009000002000000747") == [])
    check("长机型 / 3 字母版本号 / 扩展码不以机型开头只是 ⚠",
          [s for s, _ in check_rows([dict(ok_row, model="HU50A6109FUWA(1001)", ext="HU50A6109FUWA1001114",
                                         src="0100AWP0224", tgt="0100AWQ1112", change="P0224_to_Q1112")])] == []
          and [s for s, _ in check_rows([dict(ok_row, ext="75E72LIA0000697")])] == ["warn"])
    sh = check_rows([dict(ok_row, model="0100AWP0224", ext="HU58A6109FUWA", src="", tgt="")])
    check("列错位：机型是版本号、源 / 目标版本空 → ✗",
          any("机型『0100AWP0224』是版本号" in t for _, t in sh) and sum(s == "fail" for s, _ in sh) >= 3)
    check("MD5 / 大小 / 地址 / 品牌", {t.split("：")[1][:3] for s, t in check_rows(
        [dict(ok_row, md5="abc", size="0", url="http://x/a.txt", brand="123")]) if s == "fail"} >= {"MD5", "包大小", "升级包", "品牌『"})
    check("设备类型代码 / 特征码格式", len(check_rows([ok_row], "861003009000002000000747", "MT8195")) == 2)
    check("同名的两个不同升级包（目录不同也算）→ ✗",
          any("同名" in t for _, t in check_rows([ok_row, dict(ok_row, ext="32P1S9A889", md5="0" * 32,
                                                               url="ftp://u@h/b/X.zip")])))
    check("扩展码比较忽略横杠：AWOL-TTV-2150(01) / AWOL-TTV-215001077 不提醒",
          check_rows([dict(ok_row, model="AWOL-TTV-2150(01)", ext="AWOL-TTV-215001077")]) == [])
    check("共用扩展码、不同机型不算重复", not any("重复" in t for _, t in check_rows(
        [dict(ok_row, model="32DM66D", ext="DM66DEXT"), dict(ok_row, model="40DM66D", ext="DM66DEXT")])))
    check("版本变化只差末尾字母（Q1116t）→ ⚠", [s for s, _ in check_rows([dict(ok_row, change="Q0903_to_Q1116t")])] == ["warn"])
    check("品牌 / 区域为空 → ⚠（= 不限）", sorted(s for s, _ in check_rows([dict(ok_row, brand="", region="")])) == ["warn", "warn"])
    check("扩展码 + 源版本重复 → ✗", any("重复" in t for _, t in check_rows([ok_row, dict(ok_row)])))
    check("同扩展码不同源版本不算重复", not any("重复" in t for _, t in check_rows(
        [ok_row, dict(ok_row, src="0001CP1021", change="P1021_to_Q1116")])))
    check("一批两个设备类型代码 → ✗", any("2 个不同的设备类型代码" in t for _, t in check_rows(
        [dict(ok_row, code="MT8195Andr32P1S"), dict(ok_row, ext="X1", code="MT9666Android")])))
    check("源字段名转换", from_source([{"机型信息": "a", "机器扩展码": "b", "_row_number": 2}]) == [{"model": "a", "ext": "b"}])
    print("\n自测结果：", "全部通过 ✅" if good else "存在失败 ❌")
    return 0 if good else 1


if __name__ == "__main__":
    import sys
    sys.exit(_selftest() if "--selftest" in sys.argv[1:] else 0)
