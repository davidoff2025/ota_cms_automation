"""按 Rule ID 找设备类型代码（online / offline 用；--device 只是最后手段）。

查找顺序：
  1) --device（给了就全部用它）
  2) 部署库（depdb，优先有效批次）
  3) prep 最近生成的 Excel（upload 已把 Rule ID 回写进去；刚 upload、还没入库的新规则靠这一步）
  4) CMS「在线规则」页：批量规则ID 搜索，读『设备类型』列（不需要先选设备类型；只能查到已在线的规则）
     ——要开浏览器，所以放在登录后由调用方执行（from_cms_online_rules）。

实测（2026-11-12）：「规则上下线编排」的『待选规则』区不先选设备类型就搜不出结果，
所以还没上线、又不在部署库/Excel 里的规则，CMS 侧查不到设备类型代码，只能用 --device。
"""
from __future__ import annotations

from . import depdb


def resolve_local(ids: list[str], override: str = "") -> tuple[dict, dict, list[str]]:
    """不开浏览器的查找。返回 ({rid: 设备类型代码}, {rid: 来源}, [仍未找到的 rid])。"""
    if override:
        return {r: override for r in ids}, {r: "--device" for r in ids}, []
    from .active_offline import device_codes_for
    code_of, missing = device_codes_for(ids, depdb.load_all())
    src = {r: "部署库" for r in code_of}
    if missing:
        found = from_last_prep(missing)
        code_of.update(found)
        src.update({r: "prep Excel" for r in found})
        missing = [r for r in missing if r not in found]
    return code_of, src, missing


def from_last_prep(ids: list[str]) -> dict:
    """在 prep 最近生成的 Excel 里按 Rule ID 找设备类型代码。读不到返回 {}。"""
    import os
    from . import last_prep
    rec = last_prep.load()
    path = (rec or {}).get("excel", "")
    if not path or not os.path.isfile(path):
        return {}
    try:
        from .active_ota import read_import_excel
        from .dass_fill import norm_ruleid
        want = set(ids)
        out: dict[str, str] = {}
        for row in read_import_excel(path).get("rows", []):
            rid = norm_ruleid(row.get("ruleid", ""))
            dev = str(row.get("设备类型代码", "") or "").strip()
            if rid in want and dev:
                out.setdefault(rid, dev)
        return out
    except Exception:
        return {}


def from_cms_online_rules(page, ids: list[str]) -> dict:
    """在 CMS「在线规则」页按批量规则ID 查设备类型代码（只能查到已在线的规则）。页面会被导航走。"""
    if not ids:
        return {}
    from .cms_grey import chunks, open_online_rules, search_ids
    open_online_rules(page)
    out: dict[str, str] = {}
    for part in chunks(ids):
        for r in search_ids(page, part):
            rid, dev = str(r.get("规则ID", "")).strip(), str(r.get("设备类型", "")).strip()
            if rid in part and dev:
                out[rid] = dev
    return out


def print_sources(code_of: dict, src: dict, missing: list[str], color, warn) -> None:
    """按设备类型分组打印（带来源），并列出还没找到的。"""
    from .active_offline import group_by_device
    for dev, rids in group_by_device(code_of).items():
        srcs = sorted({src.get(r, "") for r in rids})
        print(f"  设备类型代码 {dev}：{', '.join(rids)}  （来自 {'、'.join(srcs)}）")
    if missing:
        print(warn(f"部署库和 prep Excel 里都没有：{', '.join(missing)} —— 将到 CMS『在线规则』页按 Rule ID 查"))
