from __future__ import annotations

import asyncio
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Literal
import json
import threading
import time
import urllib.request
import urllib.parse
import urllib.error

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse
from fastapi.responses import FileResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field


class CalcRequest(BaseModel):
    history_dates: list[str] = Field(default_factory=list)
    future_days: int = Field(ge=1)
    selected_offsets: list[int] = Field(default_factory=list)
    mode: Literal["even", "centered_even", "centered", "head", "tail", "custom"] = "even"
    preferred_start: int | None = None
    preferred_end: int | None = None
    recommend: bool = False


class DayResult(BaseModel):
    offset: int
    date: str
    selected: bool
    recommended: bool
    available: bool
    violation: bool
    past90_count: int
    remaining_after_select: int


class CalcResponse(BaseModel):
    days: list[DayResult]
    selected_offsets: list[int]
    recommended_offsets: list[int]
    violation_count: int
    message: str


class SharedStateRequest(BaseModel):
    records: list[dict] = Field(default_factory=list)
    product_sim: list[dict] = Field(default_factory=list)
    future_days: int = 90
    mode: str = "even"
    custom_start: int = 0
    custom_end: int = 14
    selected_key: str = ""
    file_name: str = ""


class SharedStateResponse(BaseModel):
    records: list[dict] = Field(default_factory=list)
    product_sim: list[dict] = Field(default_factory=list)
    future_days: int = 90
    mode: str = "even"
    custom_start: int = 0
    custom_end: int = 14
    selected_key: str = ""
    file_name: str = ""


class FeishuConfigRequest(BaseModel):
    enabled: bool = False
    app_id: str = ""
    app_secret: str = ""
    open_ids: list[str] = Field(default_factory=list)


class FeishuConfigResponse(BaseModel):
    enabled: bool = False
    app_id: str = ""
    open_ids: list[str] = Field(default_factory=list)
    has_secret: bool = False


class ReminderRange(BaseModel):
    start_date: str
    end_date: str
    asin: str = ""
    sku: str = ""
    label: str = ""


class ReminderDay(BaseModel):
    remind_date: str
    promo_start: str
    promo_end: str
    items: list[ReminderRange] = Field(default_factory=list)


app = FastAPI(title="Discount Day Calculator API")
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)
BASE_DIR = Path(__file__).resolve().parent
STATE_FILE = BASE_DIR / "shared_state.json"
DATA_DIR = BASE_DIR / "data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
SOURCE_FILE_PATH = DATA_DIR / "source_import.xlsx"
FEISHU_FILE = DATA_DIR / "feishu_config.json"
REMINDER_LOG_FILE = DATA_DIR / "reminder_log.json"
_reminder_lock = threading.Lock()


@app.get("/", response_class=HTMLResponse)
async def home() -> HTMLResponse:
    html = (BASE_DIR / "index.html").read_text(encoding="utf-8")
    return HTMLResponse(content=html)


def _load_shared_state() -> dict:
    if not STATE_FILE.exists():
        return {}
    try:
        return json.loads(STATE_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_shared_state(data: dict) -> None:
    STATE_FILE.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")


def _load_json_file(path: Path) -> dict:
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _save_json_file(path: Path, data: dict) -> None:
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def _parse_date(s: str) -> date:
    return datetime.strptime(s, "%Y-%m-%d").date()


def _build_dates(today: date, n: int) -> list[date]:
    return [today + timedelta(days=i) for i in range(n)]


def _compute(
    history_dates: list[str],
    future_days: int,
    selected_offsets: list[int],
    mode: str,
    preferred_start: int | None,
    preferred_end: int | None,
    recommend: bool,
) -> CalcResponse:
    today = date.today()
    future_dates = _build_dates(today, future_days)
    history_set = {_parse_date(x) for x in history_dates}

    # history contribution for each future day, in window [day-89, day], history only counts before today
    hist_count = [0] * future_days
    for i, d in enumerate(future_dates):
        c = 0
        for k in range(90):
            t = d - timedelta(days=k)
            if t < today and t in history_set:
                c += 1
        hist_count[i] = c

    selected = [False] * future_days
    for x in selected_offsets:
        if 0 <= x < future_days:
            selected[x] = True

    def selected_window_counts(sel: list[bool]) -> list[int]:
        pref = [0] * (future_days + 1)
        for i in range(future_days):
            pref[i + 1] = pref[i] + (1 if sel[i] else 0)
        out = [0] * future_days
        for i in range(future_days):
            l = 0 if i < 89 else i - 89
            out[i] = pref[i + 1] - pref[l]
        return out

    def validate(sel: list[bool]) -> tuple[bool, list[bool]]:
        sw = selected_window_counts(sel)
        viol = [False] * future_days
        for i in range(future_days):
            if sel[i] and (hist_count[i] + sw[i] >= 45):
                viol[i] = True
        return (not any(viol), viol)

    def can_select(idx: int, sel: list[bool]) -> bool:
        if sel[idx]:
            return True
        trial = sel[:]
        trial[idx] = True
        ok, _ = validate(trial)
        return ok

    def apply_block(start: int, seed: list[bool]) -> list[bool]:
        out = seed[:]
        for i in range(start, start + 7):
            out[i] = True
        return out

    def can_select_block(start: int, seed: list[bool], chosen_starts: list[int] | None = None) -> bool:
        if future_days < 7 or start < 0 or start + 6 >= future_days:
            return False
        if chosen_starts:
            for s in chosen_starts:
                if abs(start - s) < 7:
                    return False
        trial = apply_block(start, seed)
        ok, _ = validate(trial)
        return ok

    def recommend_even(sel_seed: list[bool]) -> list[bool]:
        # 严格7天连续 + 全区间分散：每个促销段必须连续7天，段与段不重叠
        sel = sel_seed[:]
        n = future_days
        if n < 7:
            return sel

        block_starts = list(range(0, n - 6))

        def can_select_block(start: int, seed: list[bool], chosen_starts: list[int]) -> bool:
            # 防止7天段重叠，避免“挤在一堆”
            for s in chosen_starts:
                if abs(start - s) < 7:
                    return False
            trial = seed[:]
            for i in range(start, start + 7):
                trial[i] = True
            ok, _ = validate(trial)
            return ok

        def apply_block(start: int, seed: list[bool]) -> list[bool]:
            out = seed[:]
            for i in range(start, start + 7):
                out[i] = True
            return out

        # 先估计最多可放几个7天段
        max_blocks = 0
        probe = sel[:]
        probe_starts: list[int] = []
        for s in block_starts:
            if can_select_block(s, probe, probe_starts):
                probe = apply_block(s, probe)
                probe_starts.append(s)
                max_blocks += 1

        # 从最多段数向下找首个可行的“全区间均匀”布局
        for k in range(max_blocks, 0, -1):
            trial = sel[:]
            chosen_starts: list[int] = []
            ok_all = True

            # 以未来总天数做等距锚点（按段中心点分布）
            anchors = [round(((j + 0.5) * n) / k - 0.5) for j in range(k)]
            for a in anchors:
                # 把锚点映射到可能的块起点
                near = []
                for s in block_starts:
                    center = s + 3
                    near.append((abs(center - a), s))
                near.sort(key=lambda x: x[0])

                picked = -1
                for _, s in near:
                    if can_select_block(s, trial, chosen_starts):
                        picked = s
                        break
                if picked == -1:
                    ok_all = False
                    break
                chosen_starts.append(picked)
                trial = apply_block(picked, trial)

            if ok_all:
                return trial

        return sel

    def recommend_centered_even(sel_seed: list[bool]) -> list[bool]:
        # 中心均匀也遵守7天连续块
        sel = sel_seed[:]
        n = future_days
        if n < 7:
            return sel
        starts = list(range(0, n - 6))
        center = (n - 1) / 2
        chosen_starts: list[int] = []
        while True:
            best = -1
            best_score = -10**9
            for s in starts:
                if not can_select_block(s, sel, chosen_starts):
                    continue
                block_center = s + 3
                # 兼顾分散和中心
                if chosen_starts:
                    min_dist = min(abs(s - x) for x in chosen_starts)
                else:
                    min_dist = n
                score = min_dist * 100 - abs(block_center - center)
                if score > best_score:
                    best_score = score
                    best = s
            if best == -1:
                break
            chosen_starts.append(best)
            sel = apply_block(best, sel)
        return sel

    def recommend_centered(sel_seed: list[bool]) -> list[bool]:
        sel = sel_seed[:]
        n = future_days
        center = (n - 1) / 2
        order = sorted(range(n), key=lambda i: (abs(i - center), i))
        for i in order:
            if sel[i]:
                continue
            if can_select(i, sel):
                sel[i] = True
        return sel

    def recommend_ordered(sel_seed: list[bool], order: list[int]) -> list[bool]:
        sel = sel_seed[:]
        for i in order:
            if i < 0 or i >= future_days:
                continue
            if sel[i]:
                continue
            if can_select(i, sel):
                sel[i] = True
        return sel

    def recommend_even_with_priority(sel_seed: list[bool], preferred: set[int]) -> list[bool]:
        # 自定义区间均匀分布：也按7天连续块
        sel = sel_seed[:]
        n = future_days
        if n < 7:
            return sel

        all_starts = list(range(0, n - 6))
        pref_starts = [s for s in all_starts if all((s + k) in preferred for k in range(7))]
        chosen_starts: list[int] = []

        # Phase 1: 自定义区间内尽量全选（按区间内均匀）
        while True:
            best = -1
            best_score = -10**9
            if pref_starts:
                center = (pref_starts[0] + pref_starts[-1]) / 2
            else:
                center = 0
            for s in pref_starts:
                if not can_select_block(s, sel, chosen_starts):
                    continue
                if chosen_starts:
                    min_dist = min(abs(s - x) for x in chosen_starts)
                else:
                    min_dist = n
                score = min_dist * 100 - abs(s - center)
                if score > best_score:
                    best_score = score
                    best = s
            if best == -1:
                break
            chosen_starts.append(best)
            sel = apply_block(best, sel)

        # Phase 2: 剩余块在全局均匀分布
        while True:
            best = -1
            best_score = -10**9
            global_center = (n - 1) / 2
            for s in all_starts:
                if not can_select_block(s, sel, chosen_starts):
                    continue
                block_center = s + 3
                if chosen_starts:
                    min_dist = min(abs(s - x) for x in chosen_starts)
                else:
                    min_dist = n
                score = min_dist * 100 - abs(block_center - global_center)
                if score > best_score:
                    best_score = score
                    best = s
            if best == -1:
                break
            chosen_starts.append(best)
            sel = apply_block(best, sel)
        return sel

    recommended_mask = [False] * future_days
    if recommend:
        if mode == "centered_even":
            selected = recommend_centered_even(selected)
        elif mode == "centered":
            selected = recommend_centered(selected)
        elif mode == "head":
            selected = recommend_ordered(selected, list(range(future_days)))
        elif mode == "tail":
            selected = recommend_ordered(selected, list(range(future_days - 1, -1, -1)))
        elif mode == "custom":
            s = 0 if preferred_start is None else max(0, min(future_days - 1, preferred_start))
            e = future_days - 1 if preferred_end is None else max(0, min(future_days - 1, preferred_end))
            if s <= e:
                pref = set(range(s, e + 1))
            else:
                pref = set(range(e, s + 1))
            selected = recommend_even_with_priority(selected, pref)
        else:
            selected = recommend_even(selected)
        recommended_mask = selected[:]

    ok, violations = validate(selected)
    sw = selected_window_counts(selected)

    availability = [False] * future_days
    for i in range(future_days):
        availability[i] = selected[i] or can_select(i, selected)

    days: list[DayResult] = []
    for i, d in enumerate(future_dates):
        used = hist_count[i] + sw[i]
        days.append(
            DayResult(
                offset=i,
                date=d.isoformat(),
                selected=selected[i],
                recommended=recommended_mask[i],
                available=availability[i],
                violation=violations[i],
                past90_count=used,
                remaining_after_select=max(0, 44 - used),
            )
        )

    violation_count = sum(1 for x in violations if x)
    if ok:
        msg = "当前方案满足 90 天窗口约束。"
    else:
        msg = f"检测到 {violation_count} 个已选日期不满足“该日前90天<45天”，请调整。"

    return CalcResponse(
        days=days,
        selected_offsets=[i for i, v in enumerate(selected) if v],
        recommended_offsets=[i for i, v in enumerate(recommended_mask) if v],
        violation_count=violation_count,
        message=msg,
    )


def _group_contiguous(offsets: list[int]) -> list[tuple[int, int]]:
    if not offsets:
        return []
    arr = sorted(set(int(x) for x in offsets if isinstance(x, int) or (isinstance(x, str) and x.isdigit())))
    if not arr:
        return []
    out: list[tuple[int, int]] = []
    s = arr[0]
    e = arr[0]
    for x in arr[1:]:
        if x == e + 1:
            e = x
        else:
            out.append((s, e))
            s = x
            e = x
    out.append((s, e))
    return out


def _shift_weekend_to_friday(d: date) -> date:
    w = d.weekday()
    if w == 5:
        return d - timedelta(days=1)
    if w == 6:
        return d - timedelta(days=2)
    return d


def _build_reminder_days(shared: dict, today: date | None = None) -> list[dict]:
    if today is None:
        today = date.today()
    records = {str(x.get("key", "")): x for x in shared.get("records", []) if isinstance(x, dict)}
    grouped: dict[str, dict] = {}
    for sim in shared.get("product_sim", []):
        if not isinstance(sim, dict):
            continue
        key = str(sim.get("key", "") or "")
        if not key:
            continue
        offsets = sim.get("selectedOffsets", []) or []
        ranges = _group_contiguous(offsets)
        rec = records.get(key, {})
        meta = rec.get("meta", {}) if isinstance(rec.get("meta"), dict) else {}
        label = str(rec.get("label", key) or key)
        asin = str(meta.get("ASIN", "") or "")
        sku = str(meta.get("SKU", "") or "")
        for s, e in ranges:
            if s < 0:
                continue
            promo_start = today + timedelta(days=s)
            promo_end = today + timedelta(days=e)
            remind_date = _shift_weekend_to_friday(promo_start - timedelta(days=2))
            k = remind_date.isoformat()
            if k not in grouped:
                grouped[k] = {
                    "remind_date": k,
                    "promo_start": promo_start.isoformat(),
                    "promo_end": promo_end.isoformat(),
                    "items": [],
                }
            grouped[k]["items"].append(
                {
                    "start_date": promo_start.isoformat(),
                    "end_date": promo_end.isoformat(),
                    "asin": asin,
                    "sku": sku,
                    "label": label,
                }
            )
    def _sort_key(x: dict) -> tuple[str, str]:
        return (x["remind_date"], x["promo_start"])
    return sorted(grouped.values(), key=_sort_key)


def _feishu_get_token(app_id: str, app_secret: str) -> str:
    body = json.dumps({"app_id": app_id, "app_secret": app_secret}, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(
        "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal",
        data=body,
        headers={"Content-Type": "application/json; charset=utf-8"},
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        raw = resp.read().decode("utf-8")
    data = json.loads(raw)
    if int(data.get("code", -1)) != 0:
        raise RuntimeError(f"飞书取 token 失败: {data.get('msg', 'unknown error')}")
    token = str(data.get("tenant_access_token", "") or "")
    if not token:
        raise RuntimeError("飞书 tenant_access_token 为空")
    return token


def _send_feishu_message(token: str, open_id: str, msg_type: str, content_obj: dict) -> None:
    payload = json.dumps(
        {
            "receive_id": open_id,
            "msg_type": msg_type,
            "content": json.dumps(content_obj, ensure_ascii=False),
        },
        ensure_ascii=False,
    ).encode("utf-8")
    req = urllib.request.Request(
        "https://open.feishu.cn/open-apis/im/v1/messages?receive_id_type=open_id",
        data=payload,
        headers={
            "Content-Type": "application/json; charset=utf-8",
            "Authorization": f"Bearer {token}",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=20) as resp:
        raw = resp.read().decode("utf-8")
    data = json.loads(raw)
    if int(data.get("code", -1)) != 0:
        raise RuntimeError(f"飞书发消息失败: {data.get('msg', 'unknown error')}")


def _feishu_get_json(token: str, url: str) -> dict:
    req = urllib.request.Request(
        url,
        headers={"Authorization": f"Bearer {token}"},
        method="GET",
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as resp:
            raw = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        raw = ""
        try:
            raw = e.read().decode("utf-8")
        except Exception:
            raw = ""
        if raw:
            try:
                err = json.loads(raw)
                raise RuntimeError(f"飞书接口失败 code={err.get('code')} msg={err.get('msg')} detail={err}")
            except Exception:
                raise RuntimeError(f"飞书接口失败 HTTP {e.code}: {raw}")
        raise RuntimeError(f"飞书接口失败 HTTP {e.code}")
    data = json.loads(raw)
    if int(data.get("code", -1)) != 0:
        raise RuntimeError(f"飞书接口失败 code={data.get('code')} msg={data.get('msg')} detail={data}")
    return data.get("data", {}) or {}


def _feishu_get_user_profile(token: str, open_id: str) -> dict:
    url = f"https://open.feishu.cn/open-apis/contact/v3/users/{urllib.parse.quote(open_id)}?user_id_type=open_id"
    data = _feishu_get_json(token, url)
    user = data.get("user", {}) if isinstance(data, dict) else {}
    return {
        "open_id": open_id,
        "name": str(user.get("name", "") or user.get("en_name", "") or user.get("nickname", "") or f"成员-{str(user.get('user_id', ''))[-4:]}"),
        "avatar": str(user.get("avatar", {}).get("avatar_72", "") if isinstance(user.get("avatar"), dict) else ""),
        "user_id": str(user.get("user_id", "") or ""),
    }


def _feishu_get_users_batch(token: str, open_ids: list[str]) -> list[dict]:
    params: list[tuple[str, str]] = [
        ("user_id_type", "open_id"),
        ("department_id_type", "open_department_id"),
    ]
    for oid in open_ids:
        params.append(("user_ids", oid))
    url = "https://open.feishu.cn/open-apis/contact/v3/users/batch?" + urllib.parse.urlencode(params)
    data = _feishu_get_json(token, url)
    items = data.get("items", []) if isinstance(data, dict) else []
    out: list[dict] = []
    for it in items:
        avatar_obj = it.get("avatar", {}) if isinstance(it.get("avatar"), dict) else {}
        avatar = str(
            avatar_obj.get("avatar_72")
            or avatar_obj.get("avatar_240")
            or avatar_obj.get("avatar_origin")
            or it.get("avatar_url")
            or ""
        )
        out.append(
            {
                "open_id": str(it.get("open_id", "") or ""),
                "name": str(
                    it.get("name", "")
                    or it.get("nickname", "")
                    or it.get("en_name", "")
                    or f"成员-{str(it.get('user_id', ''))[-4:]}"
                ),
                "avatar": avatar,
                "user_id": str(it.get("user_id", "") or ""),
            }
        )
    return out


def _build_message(day_row: dict) -> str:
    lines = [
        f"折扣提醒：{day_row['promo_start']} - {day_row['promo_end']} 需要添加折扣",
        f"提醒日期：{day_row['remind_date']}",
        "产品清单：",
    ]
    seen = set()
    for item in day_row.get("items", []):
        asin = str(item.get("asin", "") or "").strip()
        sku = str(item.get("sku", "") or "").strip()
        label = str(item.get("label", "") or "").strip()
        start = str(item.get("start_date", "") or "")
        end = str(item.get("end_date", "") or "")
        base = label if label else f"{asin} | {sku}".strip(" |")
        base = base if base else "-"
        key = (base, start, end)
        if key in seen:
            continue
        seen.add(key)
        lines.append(f"- {base} | {start}~{end}")
    return "\n".join(lines)


def _build_card_content(day_row: dict) -> dict:
    seen = set()
    item_lines: list[str] = []
    for item in day_row.get("items", []):
        asin = str(item.get("asin", "") or "").strip()
        sku = str(item.get("sku", "") or "").strip()
        label = str(item.get("label", "") or "").strip()
        start = str(item.get("start_date", "") or "")
        end = str(item.get("end_date", "") or "")
        base = label if label else f"{asin} | {sku}".strip(" |")
        base = base if base else "-"
        key = (base, start, end)
        if key in seen:
            continue
        seen.add(key)
        item_lines.append(f"{base} ｜ {start}~{end}")
    text = "\n".join(item_lines) if item_lines else "-"
    return {
        "config": {"wide_screen_mode": True},
        "header": {
            "template": "blue",
            "title": {"tag": "plain_text", "content": f"折扣提醒：{day_row.get('promo_start')} ~ {day_row.get('promo_end')}"},
        },
        "elements": [
            {"tag": "markdown", "content": f"**提醒日期**：{day_row.get('remind_date')}"},
            {"tag": "markdown", "content": f"**产品清单**：\n{text}"},
        ],
    }


def _send_due_reminders(force_today: str | None = None) -> dict:
    with _reminder_lock:
        cfg = _load_json_file(FEISHU_FILE)
        if not cfg.get("enabled"):
            return {"ok": True, "sent": 0, "skipped": 0, "message": "feishu disabled"}
        app_id = str(cfg.get("app_id", "") or "")
        app_secret = str(cfg.get("app_secret", "") or "")
        open_ids = [str(x).strip() for x in cfg.get("open_ids", []) if str(x).strip()]
        if not app_id or not app_secret or not open_ids:
            return {"ok": False, "sent": 0, "skipped": 0, "message": "feishu config incomplete"}
        shared = _load_shared_state()
        rows = _build_reminder_days(shared)
        today_key = force_today or date.today().isoformat()
        due = [r for r in rows if r.get("remind_date") == today_key]
        if not due:
            return {"ok": True, "sent": 0, "skipped": 0, "message": "no due reminder"}
        log = _load_json_file(REMINDER_LOG_FILE)
        sent_set = set(log.get("sent_keys", []) if isinstance(log.get("sent_keys"), list) else [])
        token = _feishu_get_token(app_id, app_secret)
        sent = 0
        skipped = 0
        for row in due:
            msg = _build_message(row)
            card = _build_card_content(row)
            for oid in open_ids:
                uniq = f"{today_key}|{oid}|{row.get('promo_start','')}|{row.get('promo_end','')}"
                if uniq in sent_set:
                    skipped += 1
                    continue
                try:
                    _send_feishu_message(token, oid, "interactive", card)
                except Exception:
                    _send_feishu_message(token, oid, "text", {"text": msg})
                sent_set.add(uniq)
                sent += 1
        _save_json_file(REMINDER_LOG_FILE, {"sent_keys": sorted(sent_set)})
        return {"ok": True, "sent": sent, "skipped": skipped, "message": "done"}


def _reminder_worker() -> None:
    while True:
        try:
            _send_due_reminders()
        except Exception:
            pass
        time.sleep(1800)


@app.post("/api/calculate", response_model=CalcResponse)
async def calculate(req: CalcRequest) -> CalcResponse:
    return await asyncio.to_thread(
        _compute,
        req.history_dates,
        req.future_days,
        req.selected_offsets,
        req.mode,
        req.preferred_start,
        req.preferred_end,
        req.recommend,
    )


@app.get("/api/shared-state", response_model=SharedStateResponse)
async def get_shared_state() -> SharedStateResponse:
    data = await asyncio.to_thread(_load_shared_state)
    return SharedStateResponse(
        records=data.get("records", []),
        product_sim=data.get("product_sim", []),
        future_days=int(data.get("future_days", 90) or 90),
        mode=str(data.get("mode", "even") or "even"),
        custom_start=int(data.get("custom_start", 0) or 0),
        custom_end=int(data.get("custom_end", 14) or 14),
        selected_key=str(data.get("selected_key", "") or ""),
        file_name=str(data.get("file_name", "") or ""),
    )


@app.post("/api/shared-state")
async def set_shared_state(req: SharedStateRequest) -> dict:
    payload = req.model_dump()
    await asyncio.to_thread(_save_shared_state, payload)
    return {"ok": True}


@app.post("/api/upload-source")
async def upload_source(file: UploadFile = File(...)) -> dict:
    content = await file.read()
    await asyncio.to_thread(SOURCE_FILE_PATH.write_bytes, content)
    data = await asyncio.to_thread(_load_shared_state)
    data["file_name"] = file.filename or "source_import.xlsx"
    await asyncio.to_thread(_save_shared_state, data)
    return {"ok": True, "file_name": data["file_name"]}


@app.get("/api/source-file")
async def download_source_file() -> FileResponse:
    if not SOURCE_FILE_PATH.exists():
        raise HTTPException(status_code=404, detail="source file not found")
    data = await asyncio.to_thread(_load_shared_state)
    fname = data.get("file_name") or "source_import.xlsx"
    return FileResponse(
        path=SOURCE_FILE_PATH,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=fname,
    )


@app.head("/api/source-file")
async def head_source_file() -> dict:
    if not SOURCE_FILE_PATH.exists():
        raise HTTPException(status_code=404, detail="source file not found")
    return {"ok": True}


@app.get("/api/feishu-config", response_model=FeishuConfigResponse)
async def get_feishu_config() -> FeishuConfigResponse:
    data = await asyncio.to_thread(_load_json_file, FEISHU_FILE)
    return FeishuConfigResponse(
        enabled=bool(data.get("enabled", False)),
        app_id=str(data.get("app_id", "") or ""),
        open_ids=[str(x).strip() for x in data.get("open_ids", []) if str(x).strip()],
        has_secret=bool(data.get("app_secret")),
    )


@app.post("/api/feishu-config")
async def set_feishu_config(req: FeishuConfigRequest) -> dict:
    old = await asyncio.to_thread(_load_json_file, FEISHU_FILE)
    payload = old.copy()
    payload["enabled"] = bool(req.enabled)
    payload["app_id"] = req.app_id.strip()
    payload["open_ids"] = [x.strip() for x in req.open_ids if x.strip()]
    if req.app_secret.strip():
        payload["app_secret"] = req.app_secret.strip()
    await asyncio.to_thread(_save_json_file, FEISHU_FILE, payload)
    return {"ok": True}


@app.get("/api/reminders/preview", response_model=list[ReminderDay])
async def preview_reminders() -> list[ReminderDay]:
    shared = await asyncio.to_thread(_load_shared_state)
    rows = await asyncio.to_thread(_build_reminder_days, shared)
    return [ReminderDay(**x) for x in rows]


@app.post("/api/reminders/send-now")
async def send_reminders_now() -> dict:
    return await asyncio.to_thread(_send_due_reminders)


@app.get("/api/feishu/departments")
async def feishu_departments() -> dict:
    try:
        cfg = await asyncio.to_thread(_load_json_file, FEISHU_FILE)
        app_id = str(cfg.get("app_id", "") or "")
        app_secret = str(cfg.get("app_secret", "") or "")
        if not app_id or not app_secret:
            raise HTTPException(status_code=400, detail="请先保存飞书 app_id/app_secret")
        token = await asyncio.to_thread(_feishu_get_token, app_id, app_secret)
        root_url = "https://open.feishu.cn/open-apis/contact/v3/departments/0/children?department_id_type=department_id&page_size=50"
        data = await asyncio.to_thread(_feishu_get_json, token, root_url)
        items = data.get("items", []) if isinstance(data, dict) else []
        return {"items": items}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"加载部门失败: {e}")


@app.get("/api/feishu/users")
async def feishu_users(department_id: str) -> dict:
    try:
        cfg = await asyncio.to_thread(_load_json_file, FEISHU_FILE)
        app_id = str(cfg.get("app_id", "") or "")
        app_secret = str(cfg.get("app_secret", "") or "")
        if not app_id or not app_secret:
            raise HTTPException(status_code=400, detail="请先保存飞书 app_id/app_secret")
        if not department_id.strip():
            raise HTTPException(status_code=400, detail="department_id 不能为空")
        token = await asyncio.to_thread(_feishu_get_token, app_id, app_secret)
        q = urllib.parse.urlencode(
            {
                "department_id": department_id.strip(),
                "department_id_type": "open_department_id",
                "user_id_type": "open_id",
                "page_size": 50,
            }
        )
        url = f"https://open.feishu.cn/open-apis/contact/v3/users/find_by_department?{q}"
        data = await asyncio.to_thread(_feishu_get_json, token, url)
        items = data.get("items", []) if isinstance(data, dict) else []
        return {"items": items}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"加载成员失败: {e}")


@app.on_event("startup")
async def startup_event() -> None:
    th = threading.Thread(target=_reminder_worker, daemon=True, name="feishu-reminder-worker")
    th.start()
@app.get("/api/feishu/scope-users")
async def feishu_scope_users() -> dict:
    try:
        cfg = await asyncio.to_thread(_load_json_file, FEISHU_FILE)
        app_id = str(cfg.get("app_id", "") or "")
        app_secret = str(cfg.get("app_secret", "") or "")
        if not app_id or not app_secret:
            raise HTTPException(status_code=400, detail="请先保存飞书 app_id/app_secret")
        token = await asyncio.to_thread(_feishu_get_token, app_id, app_secret)
        data = await asyncio.to_thread(_feishu_get_json, token, "https://open.feishu.cn/open-apis/contact/v3/scopes")
        user_ids = data.get("user_ids", []) if isinstance(data, dict) else []
        open_ids = [str(x).strip() for x in user_ids if str(x).strip()]
        items: list[dict] = []
        if open_ids:
            # 先走批量接口，速度更快，信息更完整
            for i in range(0, len(open_ids), 50):
                chunk = open_ids[i : i + 50]
                try:
                    rows = await asyncio.to_thread(_feishu_get_users_batch, token, chunk)
                    got = {x.get("open_id", ""): x for x in rows}
                    for oid in chunk:
                        if oid in got:
                            items.append(got[oid])
                        else:
                            items.append({"open_id": oid, "name": "成员", "avatar": ""})
                except Exception:
                    # 批量失败回退单查，确保不丢人
                    for oid in chunk:
                        try:
                            prof = await asyncio.to_thread(_feishu_get_user_profile, token, oid)
                            items.append(prof)
                        except Exception:
                            items.append({"open_id": oid, "name": "成员", "avatar": ""})
        return {"items": items}
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"加载成员失败: {e}")
