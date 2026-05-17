"""花费统计仪表盘 - Flask 后端 (含数据管理)"""

import csv
import glob
import json
import os
import subprocess
import sys
from collections import defaultdict
from datetime import datetime, timezone
from flask import Flask, jsonify, render_template, request

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ALIASES_FILE = os.path.join(BASE_DIR, "user-aliases.json")
DELETED_USERS_FILE = os.path.join(BASE_DIR, "deleted-users.json")
FETCH_STATUS_FILE = os.path.join(BASE_DIR, ".fetch-status.json")

# ── 别名管理 ──────────────────────────────────────────────

def load_aliases():
    """加载用户名 → 备注映射"""
    if os.path.exists(ALIASES_FILE):
        with open(ALIASES_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}


def save_aliases(aliases):
    """保存别名到 JSON 文件"""
    with open(ALIASES_FILE, "w", encoding="utf-8") as f:
        json.dump(aliases, f, ensure_ascii=False, indent=2)


def load_deleted_users():
    """加载已删除用户列表"""
    if os.path.exists(DELETED_USERS_FILE):
        with open(DELETED_USERS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def save_deleted_users(users):
    """保存已删除用户列表"""
    with open(DELETED_USERS_FILE, "w", encoding="utf-8") as f:
        json.dump(users, f, ensure_ascii=False, indent=2)


def apply_alias(name):
    """如果存在备注则返回备注，否则返回原名"""
    aliases = load_aliases()
    return aliases.get(name, name)


# ── 数据加载 (带文件路径和行号) ──────────────────────────

def load_all_amount_files():
    """加载所有 amount-*.csv 文件，每行附加 _file 和 _idx，过滤已删除用户"""
    deleted = set(load_deleted_users())
    rows = []
    pattern = os.path.join(BASE_DIR, "amount-*-*.csv")
    for filepath in sorted(glob.glob(pattern)):
        fname = os.path.basename(filepath)
        with open(filepath, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for idx, row in enumerate(reader):
                if row.get("api_key_name", "") in deleted:
                    continue
                row["_file"] = fname
                row["_idx"] = idx
                rows.append(row)
    return rows


def load_all_cost_files():
    """加载所有 cost-*.csv 文件，每行附加 _file 和 _idx"""
    rows = []
    pattern = os.path.join(BASE_DIR, "cost-*-*.csv")
    for filepath in sorted(glob.glob(pattern)):
        fname = os.path.basename(filepath)
        with open(filepath, newline="", encoding="utf-8") as f:
            reader = csv.DictReader(f)
            for idx, row in enumerate(reader):
                row["_file"] = fname
                row["_idx"] = idx
                rows.append(row)
    return rows


def get_available_dates():
    dates = set()
    for row in load_all_amount_files():
        dates.add(row["utc_date"])
    for row in load_all_cost_files():
        dates.add(row["utc_date"])
    return sorted(dates)


def filter_by_dates(rows, selected_dates):
    if not selected_dates:
        return rows
    return [r for r in rows if r["utc_date"] in selected_dates]


# ── CSV 写回辅助 ─────────────────────────────────────────

def _amount_fieldnames():
    return ["user_id", "utc_date", "model", "api_key_name", "api_key",
            "type", "price", "amount"]


def _cost_fieldnames():
    return ["user_id", "utc_date", "model", "wallet_type", "cost", "currency"]


def _rewrite_csv(filename, fieldnames, rows):
    """用新的 rows 覆写 CSV 文件（临时文件 + 原子替换，避免 Windows 权限问题）"""
    filepath = os.path.join(BASE_DIR, filename)
    import tempfile
    tmp_fd, tmp_path = tempfile.mkstemp(dir=BASE_DIR, suffix=".csv")
    try:
        with os.fdopen(tmp_fd, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for row in rows:
                clean = {k: row.get(k, "") for k in fieldnames}
                writer.writerow(clean)
        os.replace(tmp_path, filepath)
    except BaseException:
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
        raise


# ── 页面 ──────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


# ── 别名 API ──────────────────────────────────────────────

@app.route("/api/aliases", methods=["GET", "PUT"])
def aliases():
    if request.method == "GET":
        return jsonify(load_aliases())
    else:
        data = request.get_json()
        if not isinstance(data, dict):
            return jsonify({"error": "Expected a JSON object"}), 400
        save_aliases(data)
        return jsonify({"ok": True})


# ── 数据查询 API ──────────────────────────────────────────

@app.route("/api/date-range")
def date_range():
    dates = get_available_dates()
    return jsonify({
        "dates": dates,
        "min": dates[0] if dates else None,
        "max": dates[-1] if dates else None,
    })


@app.route("/api/summary")
def summary():
    selected = request.args.get("dates", "")
    selected_dates = set(selected.split(",")) if selected else set()

    cost_rows = filter_by_dates(load_all_cost_files(), selected_dates)
    total_cost = sum(float(r["cost"]) for r in cost_rows)
    active_dates = set(r["utc_date"] for r in cost_rows)
    days = max(len(active_dates), 1)

    amount_rows = filter_by_dates(load_all_amount_files(), selected_dates)
    total_requests = sum(
        int(r["amount"]) for r in amount_rows if r["type"] == "request_count"
    )

    return jsonify({
        "total_cost": round(total_cost, 2),
        "avg_daily_cost": round(total_cost / days, 2),
        "total_requests": total_requests,
        "active_days": len(active_dates),
        "currency": "CNY",
    })


@app.route("/api/daily-cost")
def daily_cost():
    selected = request.args.get("dates", "")
    selected_dates = set(selected.split(",")) if selected else set()
    cost_rows = filter_by_dates(load_all_cost_files(), selected_dates)

    daily = defaultdict(lambda: defaultdict(float))
    models = set()
    for r in cost_rows:
        daily[r["utc_date"]][r["model"]] += float(r["cost"])
        models.add(r["model"])

    dates = sorted(daily.keys())
    models = sorted(models)
    datasets = []
    for model in models:
        datasets.append({
            "label": model,
            "data": [round(daily[d].get(model, 0), 4) for d in dates],
        })
    return jsonify({"dates": dates, "datasets": datasets})


@app.route("/api/cost-by-model")
def cost_by_model():
    selected = request.args.get("dates", "")
    selected_dates = set(selected.split(",")) if selected else set()
    cost_rows = filter_by_dates(load_all_cost_files(), selected_dates)
    model_cost = defaultdict(float)
    for r in cost_rows:
        model_cost[r["model"]] += float(r["cost"])
    sorted_items = sorted(model_cost.items(), key=lambda x: x[1], reverse=True)
    return jsonify({
        "labels": [x[0] for x in sorted_items],
        "data": [round(x[1], 2) for x in sorted_items],
    })


@app.route("/api/cost-by-user")
def cost_by_user():
    selected = request.args.get("dates", "")
    selected_dates = set(selected.split(",")) if selected else set()
    amount_rows = filter_by_dates(load_all_amount_files(), selected_dates)

    user_cost = defaultdict(float)
    for r in amount_rows:
        if r["type"] == "request_count":
            continue
        price = float(r["price"]) if r["price"] else 0
        amt = float(r["amount"]) if r["amount"] else 0
        user_cost[r["api_key_name"]] += price * amt

    sorted_items = sorted(user_cost.items(), key=lambda x: x[1], reverse=True)
    return jsonify({
        "labels": [apply_alias(x[0]) for x in sorted_items],
        "data": [round(x[1], 2) for x in sorted_items],
    })


@app.route("/api/token-usage")
def token_usage():
    selected = request.args.get("dates", "")
    selected_dates = set(selected.split(",")) if selected else set()
    amount_rows = filter_by_dates(load_all_amount_files(), selected_dates)

    models = sorted(set(r["model"] for r in amount_rows))
    token_types = ["output_tokens", "input_cache_hit_tokens", "input_cache_miss_tokens"]

    datasets = []
    for tt in token_types:
        data = []
        for model in models:
            total = sum(
                int(r["amount"])
                for r in amount_rows
                if r["model"] == model and r["type"] == tt
            )
            data.append(total)
        datasets.append({
            "label": tt.replace("_", " ").title(),
            "data": data,
        })
    return jsonify({"models": models, "datasets": datasets})


# ── 数据管理 API（编辑/删除）──────────────────────────────

@app.route("/api/table-data")
def table_data():
    """返回所有数据的扁平表格视图，附加别名"""
    aliases_map = load_aliases()

    amount_rows = load_all_amount_files()
    cost_rows = load_all_cost_files()

    result = []
    for r in amount_rows:
        name = r.get("api_key_name", "")
        cost_val = 0.0
        if r["type"] != "request_count":
            price = float(r["price"]) if r["price"] else 0
            amt = float(r["amount"]) if r["amount"] else 0
            cost_val = round(price * amt, 6)

        result.append({
            "source": "amount",
            "file": r["_file"],
            "idx": r["_idx"],
            "utc_date": r["utc_date"],
            "model": r["model"],
            "api_key_name": name,
            "api_key_name_alias": aliases_map.get(name, name),
            "type": r["type"],
            "price": r.get("price", ""),
            "amount": r.get("amount", ""),
            "cost": cost_val,
        })

    for r in cost_rows:
        result.append({
            "source": "cost",
            "file": r["_file"],
            "idx": r["_idx"],
            "utc_date": r["utc_date"],
            "model": r["model"],
            "api_key_name": "",
            "api_key_name_alias": "",
            "type": "daily_cost",
            "price": "",
            "amount": "",
            "cost": float(r["cost"]),
        })

    return jsonify(result)


@app.route("/api/delete-row", methods=["POST"])
def delete_row():
    data = request.get_json()
    source = data.get("source")      # "amount" | "cost"
    filename = data.get("file")      # e.g. "amount-2026-5.csv"
    idx = data.get("idx")            # row index

    if source not in ("amount", "cost") or filename is None or idx is None:
        return jsonify({"error": "Invalid parameters"}), 400

    filepath = os.path.join(BASE_DIR, filename)
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404

    fieldnames = _amount_fieldnames() if source == "amount" else _cost_fieldnames()

    # 读取所有行
    with open(filepath, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    if idx < 0 or idx >= len(rows):
        return jsonify({"error": "Index out of range"}), 400

    del rows[idx]
    _rewrite_csv(filename, fieldnames, rows)
    return jsonify({"ok": True})


@app.route("/api/edit-row", methods=["POST"])
def edit_row():
    data = request.get_json()
    source = data.get("source")      # "amount" | "cost"
    filename = data.get("file")
    idx = data.get("idx")
    values = data.get("values")      # dict of field→new_value

    if source not in ("amount", "cost") or filename is None or idx is None or not values:
        return jsonify({"error": "Invalid parameters"}), 400

    filepath = os.path.join(BASE_DIR, filename)
    if not os.path.exists(filepath):
        return jsonify({"error": "File not found"}), 404

    fieldnames = _amount_fieldnames() if source == "amount" else _cost_fieldnames()

    # 读取所有行
    with open(filepath, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        rows = list(reader)

    if idx < 0 or idx >= len(rows):
        return jsonify({"error": "Index out of range"}), 400

    # 只更新允许的字段
    for k, v in values.items():
        if k in fieldnames:
            rows[idx][k] = str(v)

    _rewrite_csv(filename, fieldnames, rows)
    return jsonify({"ok": True})


@app.route("/api/delete-user", methods=["POST"])
def delete_user():
    """将用户加入黑名单, 所有数据在加载时自动过滤"""
    data = request.get_json()
    api_key_name = data.get("api_key_name", "").strip()
    if not api_key_name:
        return jsonify({"error": "Missing api_key_name"}), 400

    # 统计受影响的行数 (先不过滤读取)
    count = 0
    pattern = os.path.join(BASE_DIR, "amount-*-*.csv")
    for filepath in sorted(glob.glob(pattern)):
        with open(filepath, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("api_key_name", "") == api_key_name:
                    count += 1

    # 写入黑名单
    deleted = load_deleted_users()
    if api_key_name not in deleted:
        deleted.append(api_key_name)
        save_deleted_users(deleted)

    # 清理别名
    aliases = load_aliases()
    if api_key_name in aliases:
        del aliases[api_key_name]
        save_aliases(aliases)

    return jsonify({"ok": True, "deleted": count, "user": api_key_name})


@app.route("/api/deleted-users", methods=["GET"])
def get_deleted_users():
    """获取已删除用户列表"""
    return jsonify(load_deleted_users())


@app.route("/api/restore-user", methods=["POST"])
def restore_user():
    """恢复被删除的用户"""
    data = request.get_json()
    api_key_name = data.get("api_key_name", "").strip()
    if not api_key_name:
        return jsonify({"error": "Missing api_key_name"}), 400

    deleted = load_deleted_users()
    if api_key_name in deleted:
        deleted.remove(api_key_name)
        save_deleted_users(deleted)
        return jsonify({"ok": True, "user": api_key_name})
    return jsonify({"ok": False, "error": "User not in deleted list"}), 404


@app.route("/api/delete-unalised", methods=["POST"])
def delete_unalised():
    """一键删除所有未备注用户：有备注=白名单保留，无备注=加入黑名单"""
    aliases = load_aliases()
    deleted = set(load_deleted_users())

    # 扫描所有用户
    all_users = set()
    pattern = os.path.join(BASE_DIR, "amount-*-*.csv")
    for filepath in sorted(glob.glob(pattern)):
        with open(filepath, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                name = row.get("api_key_name", "").strip()
                if name:
                    all_users.add(name)

    # 未备注的加入黑名单
    new_deleted = []
    for user in sorted(all_users):
        if user not in aliases and user not in deleted:
            new_deleted.append(user)

    # 更新黑名单
    current = load_deleted_users()
    for user in new_deleted:
        if user not in current:
            current.append(user)
    save_deleted_users(current)

    # 统计实际行数
    total_rows = 0
    for filepath in sorted(glob.glob(pattern)):
        with open(filepath, newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                if row.get("api_key_name", "") in new_deleted:
                    total_rows += 1

    return jsonify({
        "ok": True,
        "deleted_users": new_deleted,
        "deleted_rows": total_rows,
        "kept_users": [u for u in sorted(all_users) if u not in current and u not in new_deleted],
    })


# ── 在线更新 API ──────────────────────────────────────────

@app.route("/api/fetch-online", methods=["POST"])
def fetch_online():
    """启动在线数据获取（后台子进程）"""
    # 检查是否已有任务在运行
    if os.path.exists(FETCH_STATUS_FILE):
        try:
            with open(FETCH_STATUS_FILE, "r", encoding="utf-8") as f:
                current = json.load(f)
            if current.get("status") in ("running", "awaiting_login"):
                return jsonify({"ok": False, "running": True, "message": "已有任务在运行中"})
        except Exception:
            pass

    # 写入初始状态
    status = {"status": "running", "message": "正在启动浏览器...", "started_at": datetime.now(timezone.utc).isoformat()}
    tmp = FETCH_STATUS_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(status, f, ensure_ascii=False)
    os.replace(tmp, FETCH_STATUS_FILE)

    # 启动后台子进程
    script_path = os.path.join(BASE_DIR, "fetch_online.py")
    subprocess.Popen(
        [sys.executable, script_path, FETCH_STATUS_FILE],
        cwd=BASE_DIR,
        creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0,
    )

    return jsonify({"ok": True, "message": "已启动"})


@app.route("/api/fetch-status")
def fetch_status():
    """获取在线更新的当前状态"""
    if not os.path.exists(FETCH_STATUS_FILE):
        return jsonify({"status": "idle"})

    try:
        with open(FETCH_STATUS_FILE, "r", encoding="utf-8") as f:
            data = json.load(f)

        # 检查超时（10 分钟）
        if data.get("status") in ("running", "awaiting_login"):
            started = data.get("started_at", "")
            if started:
                try:
                    st = datetime.fromisoformat(started)
                    elapsed = (datetime.now(timezone.utc) - st).total_seconds()
                    if elapsed > 600:
                        return jsonify({"status": "error", "message": "任务超时（10分钟），请重试"})
                except Exception:
                    pass

        return jsonify(data)
    except Exception:
        return jsonify({"status": "idle"})


if __name__ == "__main__":
    print("Budget Dashboard: http://localhost:5000")
    app.run(debug=True, host="0.0.0.0", port=5000)
