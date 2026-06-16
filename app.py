# -*- coding: utf-8 -*-
"""Nikaido Research DB - 単一ファイル版（iPhoneデプロイ用）

constants / db / analysis / UI を1ファイルに統合。
データ保存先: Googleスプレッドシート。
認証情報は Streamlit secrets から読み込む:
  - spreadsheet_id           : スプレッドシートID
  - gcp_service_account_json : サービスアカウントJSONの中身（まるごと文字列）
"""
import json
import datetime as dt

# 日本標準時（JST = UTC+9）。サーバーがUTCで動くため、日付・時刻は必ずJSTで扱う。
JST = dt.timezone(dt.timedelta(hours=9))


def now_jst():
    return dt.datetime.now(JST)

import numpy as np
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials


# =========================================================
# 選択肢・定数
# =========================================================
SIDES = ["買い", "空売り"]

DISCOVERY_ROUTES = [
    "前日比値上がり", "前日比値下がり", "10分前比値上がり", "10分前比値下がり",
    "出来高急増", "寄り前気配", "X", "ニュース", "株ドラゴン", "監視リスト", "その他",
]
WATCHLIST_TIMINGS = ["前日夜", "寄り前", "前場", "後場"]
ENTRY_METHODS = [
    "二階堂型", "ガジャラ型", "プルバック高値突破", "VWAP反発", "フィボ50%",
    "フィボ61.8%", "ギャップダウンリバウンド", "ストップ安リバウンド", "急騰押し目", "その他",
]
CONDITIONS = [
    "急騰銘柄", "急落銘柄", "材料あり", "VWAP上", "VWAP下", "EMA上", "EMA下",
    "売り枯れ", "買い出現", "出来高減少", "出来高増加", "高値更新", "安値切り上げ",
    "フィボ50%", "フィボ61.8%", "長い下ヒゲ", "歩み値買い連打", "歩み値売り連打",
]
CRASH_PERSONALITIES = [
    "短期投げ尽くし型", "本物崩壊型", "アルゴ雪崩型", "利確連鎖型", "洗い落とし型",
    "GU失敗型", "本尊撤退型", "2波成功型", "2波失敗型", "未分類",
]
PRIOR5MIN_STATES = ["急騰中", "押し目中", "ヨコヨコ", "ブレイク直前", "ブレイク後"]

STOCK_MASTER_SEED = {"7746": "岡本硝子"}

SPECIAL_COMPARISONS = [
    ("売り枯れ あり vs なし", "売り枯れ", None),
    ("VWAP上 vs VWAP下", "VWAP上", "VWAP下"),
    ("高値更新 あり vs なし", "高値更新", None),
    ("買い出現 あり vs なし", "買い出現", None),
]


# =========================================================
# データ層（Googleスプレッドシート）
# =========================================================
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TRADES_HEADERS = [
    "id", "trade_date", "entry_time", "stock_code", "stock_name", "side",
    "shares", "in_price", "out_price", "pnl", "discovery_route",
    "watchlist_timing", "entry_methods", "conditions", "crash_personality",
    "prior5min_state", "memo", "screenshot_url", "created_at",
]
STOCK_HEADERS = ["code", "name"]


def _load_credentials():
    if "gcp_service_account_json" in st.secrets:
        info = json.loads(st.secrets["gcp_service_account_json"])
    elif "gcp_service_account" in st.secrets:
        info = dict(st.secrets["gcp_service_account"])
    else:
        raise KeyError("secrets に gcp_service_account_json がありません。")
    return Credentials.from_service_account_info(info, scopes=SCOPES)


@st.cache_resource
def _spreadsheet():
    gc = gspread.authorize(_load_credentials())
    return gc.open_by_key(st.secrets["spreadsheet_id"])


def _ws(name, headers):
    ss = _spreadsheet()
    try:
        ws = ss.worksheet(name)
    except gspread.exceptions.WorksheetNotFound:
        ws = ss.add_worksheet(title=name, rows=2000, cols=max(10, len(headers)))
        ws.append_row(headers, value_input_option="RAW")
        return ws
    if not ws.row_values(1):
        ws.append_row(headers, value_input_option="RAW")
    return ws


@st.cache_resource
def init_db():
    _ws("trades", TRADES_HEADERS)
    _ws("stock_master", STOCK_HEADERS)
    master = _load_stock_master()
    for code, name in STOCK_MASTER_SEED.items():
        if code not in master:
            upsert_stock(code, name)
    return True


def calc_pnl(side, shares, in_price, out_price):
    if None in (shares, in_price, out_price):
        return 0.0
    if side == "空売り":
        return (in_price - out_price) * shares
    return (out_price - in_price) * shares


@st.cache_data(ttl=600)
def _load_stock_master():
    ws = _ws("stock_master", STOCK_HEADERS)
    out = {}
    for r in ws.get_all_records():
        code = str(r.get("code", "")).strip()
        name = str(r.get("name", "")).strip()
        if code:
            out[code] = name
    return out


def lookup_stock(code):
    if not code:
        return None
    return _load_stock_master().get(code.strip())


def upsert_stock(code, name):
    code = (code or "").strip()
    name = (name or "").strip()
    if not code or not name:
        return
    ws = _ws("stock_master", STOCK_HEADERS)
    codes = ws.col_values(1)
    if code in codes[1:]:
        ws.update_cell(codes.index(code) + 1, 2, name)
    else:
        ws.append_row([code, name], value_input_option="RAW")
    _load_stock_master.clear()


def all_stocks_df():
    m = _load_stock_master()
    if not m:
        return pd.DataFrame(columns=["code", "name"])
    return pd.DataFrame(sorted(m.items()), columns=["code", "name"])


def import_stock_csv(df):
    cols = {c.lower(): c for c in df.columns}
    code_col = cols.get("code") or df.columns[0]
    name_col = cols.get("name") or df.columns[1]
    master = dict(_load_stock_master())
    n = 0
    for _, r in df.iterrows():
        code = str(r[code_col]).strip()
        name = str(r[name_col]).strip()
        if code and name and code.lower() != "nan":
            master[code] = name
            n += 1
    ws = _ws("stock_master", STOCK_HEADERS)
    ws.clear()
    ws.append_rows([STOCK_HEADERS] + [[k, v] for k, v in sorted(master.items())],
                   value_input_option="RAW")
    _load_stock_master.clear()
    return n


@st.cache_data(ttl=60)
def _load_trades_records():
    return _ws("trades", TRADES_HEADERS).get_all_records()


def _next_id(recs):
    ids = [int(r["id"]) for r in recs if str(r.get("id", "")).strip().isdigit()]
    return max(ids) + 1 if ids else 1


def add_trade(d):
    pnl = calc_pnl(d["side"], d["shares"], d["in_price"], d["out_price"])
    ws = _ws("trades", TRADES_HEADERS)
    new_id = _next_id(_load_trades_records())
    created = now_jst().strftime("%Y-%m-%d %H:%M:%S")
    ws.append_row([
        new_id, d["trade_date"], d["entry_time"], d["stock_code"], d["stock_name"],
        d["side"], d["shares"], d["in_price"], d["out_price"], pnl,
        d["discovery_route"], d["watchlist_timing"],
        json.dumps(d["entry_methods"], ensure_ascii=False),
        json.dumps(d["conditions"], ensure_ascii=False),
        d["crash_personality"], d["prior5min_state"], d["memo"],
        d.get("screenshot_url", ""), created,
    ], value_input_option="RAW")
    _load_trades_records.clear()
    return new_id


def delete_trade(tid):
    ws = _ws("trades", TRADES_HEADERS)
    ids = ws.col_values(1)
    target = str(int(tid))
    for i, v in enumerate(ids):
        if i == 0:
            continue
        if str(v).strip() == target:
            ws.delete_rows(i + 1)
            break
    _load_trades_records.clear()


def _safe_json(s):
    try:
        v = json.loads(s) if s else []
        return v if isinstance(v, list) else []
    except Exception:
        return []


def _hour_band(t):
    if not t or not isinstance(t, str) or ":" not in t:
        return "不明"
    try:
        return f"{int(t.split(':')[0])}時台"
    except Exception:
        return "不明"


def get_trades_df():
    recs = _load_trades_records()
    if not recs:
        return pd.DataFrame()
    df = pd.DataFrame(recs)
    for c in TRADES_HEADERS:
        if c not in df.columns:
            df[c] = None
    df["shares"] = pd.to_numeric(df["shares"], errors="coerce").fillna(0).astype(int)
    for c in ["in_price", "out_price", "pnl", "id"]:
        df[c] = pd.to_numeric(df[c], errors="coerce")
    df["id"] = df["id"].fillna(0).astype(int)
    df["entry_methods"] = df["entry_methods"].apply(_safe_json)
    df["conditions"] = df["conditions"].apply(_safe_json)
    df = df.sort_values(["trade_date", "id"], ascending=[False, False]).reset_index(drop=True)
    dt_ = pd.to_datetime(df["trade_date"], errors="coerce")
    wd = {0: "月", 1: "火", 2: "水", 3: "木", 4: "金", 5: "土", 6: "日"}
    df["曜日"] = dt_.dt.dayofweek.map(wd)
    df["時間帯"] = df["entry_time"].apply(_hour_band)
    return df


# =========================================================
# 分析
# =========================================================
def _stats(pnl):
    total = len(pnl)
    if total == 0:
        return pd.Series({"件数": 0, "勝率(%)": 0, "平均利益": 0, "平均損失": 0, "PF": 0, "期待値": 0})
    wins = pnl[pnl > 0]
    losses = pnl[pnl < 0]
    gp = wins.sum()
    gl = abs(losses.sum())
    pf = (gp / gl) if gl > 0 else (np.inf if gp > 0 else 0)
    return pd.Series({
        "件数": int(total),
        "勝率(%)": round(len(wins) / total * 100, 1),
        "平均利益": round(float(wins.mean()) if len(wins) else 0, 0),
        "平均損失": round(float(losses.mean()) if len(losses) else 0, 0),
        "PF": round(float(pf), 2) if np.isfinite(pf) else np.inf,
        "期待値": round(float(pnl.mean()), 0),
    })


def overall_summary(df):
    return _stats(df["pnl"]) if not df.empty else None


def group_ranking(df, col, min_count=1):
    if df.empty or col not in df.columns:
        return pd.DataFrame()
    g = df.groupby(col)["pnl"].apply(_stats)
    if g.empty:
        return pd.DataFrame()
    res = g.unstack()
    res = res[res["件数"] >= min_count].copy()
    if res.empty:
        return res
    res["件数"] = res["件数"].astype(int)
    return res.sort_values("期待値", ascending=False).reset_index().rename(columns={col: "項目"})


def _explode(df, list_col):
    rows = []
    for _, r in df.iterrows():
        for item in (r[list_col] or []):
            rows.append({"項目": item, "pnl": r["pnl"]})
    return pd.DataFrame(rows)


def list_ranking(df, list_col, min_count=1):
    ex = _explode(df, list_col)
    if ex.empty:
        return pd.DataFrame()
    res = ex.groupby("項目")["pnl"].apply(_stats).unstack()
    res = res[res["件数"] >= min_count].copy()
    if res.empty:
        return res
    res["件数"] = res["件数"].astype(int)
    return res.sort_values("期待値", ascending=False).reset_index()


def _has_tag(df, tag):
    return df[df["conditions"].apply(lambda lst: tag in (lst or []))]


def special_comparison(df, tag_a, tag_b=None):
    a = _has_tag(df, tag_a)
    if tag_b is None:
        b = df[~df.index.isin(a.index)]
        la, lb = f"{tag_a} あり", f"{tag_a} なし"
    else:
        b = _has_tag(df, tag_b)
        la, lb = tag_a, tag_b
    out = pd.DataFrame([_stats(a["pnl"]).rename(la), _stats(b["pnl"]).rename(lb)])
    out["件数"] = out["件数"].astype(int)
    return out.reset_index().rename(columns={"index": "区分"})


def generate_insights(df, min_count=3):
    if df.empty:
        return ["まだトレード記録がありません。データを蓄積すると分析できます。"]
    if len(df) < min_count:
        return [f"記録が{len(df)}件です。最低{min_count}件たまると傾向分析が安定します。"]

    lines = []
    ov = overall_summary(df)
    lines.append(f"📊 全体: {int(ov['件数'])}件 / 勝率{ov['勝率(%)']}% / "
                 f"期待値{int(ov['期待値']):+,}円 / PF{ov['PF']}")

    factors = []
    for col in ["discovery_route", "prior5min_state", "crash_personality"]:
        for _, row in group_ranking(df, col, min_count).iterrows():
            factors.append((row["項目"], row["期待値"], int(row["件数"]), row["勝率(%)"]))
    for col in ["entry_methods", "conditions"]:
        for _, row in list_ranking(df, col, min_count).iterrows():
            factors.append((row["項目"], row["期待値"], int(row["件数"]), row["勝率(%)"]))
    if factors:
        factors.sort(key=lambda x: x[1], reverse=True)
        b = factors[0]
        lines.append(f"✅ 単一要因で最も期待値が高いのは「{b[0]}」"
                     f"（期待値{int(b[1]):+,}円 / 勝率{b[3]}% / {b[2]}件）です。")
        w = factors[-1]
        if w[1] < 0:
            lines.append(f"⚠️ 「{w[0]}」は期待値マイナス"
                         f"（{int(w[1]):+,}円 / 勝率{w[3]}% / {w[2]}件）。単独でのエントリー根拠は注意。")

    combos = []
    methods = sorted({m for lst in df["entry_methods"] for m in lst})
    conds = sorted({c for lst in df["conditions"] for c in lst})
    for m in methods:
        sub_m = df[df["entry_methods"].apply(lambda lst: m in lst)]
        for c in conds:
            sub = sub_m[sub_m["conditions"].apply(lambda lst: c in lst)]
            if len(sub) >= min_count:
                combos.append((f"{m}＋{c}", float(sub["pnl"].mean()),
                               len(sub), round((sub["pnl"] > 0).mean() * 100, 1)))
    if combos:
        combos.sort(key=lambda x: x[1], reverse=True)
        b = combos[0]
        lines.append(f"🔥 組み合わせで最強は「{b[0]}」"
                     f"（期待値{int(b[1]):+,}円 / 勝率{b[3]}% / {b[2]}件）。")
        if combos[-1][1] < 0:
            w = combos[-1]
            lines.append(f"🧊 逆に「{w[0]}」は期待値{int(w[1]):+,}円で負けパターン。")

    dfx = df.copy()
    dfx["_dt"] = pd.to_datetime(dfx["trade_date"], errors="coerce")
    recent = dfx.sort_values(["_dt", "id"]).tail(10)
    if len(recent) >= 3:
        r_exp = recent["pnl"].mean()
        diff = r_exp - df["pnl"].mean()
        trend = "改善" if diff > 0 else ("横ばい" if diff == 0 else "悪化")
        lines.append(f"📈 直近{len(recent)}件の期待値は{int(r_exp):+,}円"
                     f"（全体比 {int(diff):+,}円・{trend}傾向）。")
    return lines


# =========================================================
# UI
# =========================================================
st.set_page_config(page_title="Nikaido Research DB", page_icon="📈", layout="centered")

try:
    init_db()
except Exception as e:
    st.error("スプレッドシートに接続できません。Secrets（gcp_service_account_json / spreadsheet_id）と、"
             "サービスアカウントへのシート共有（編集者）を確認してください。")
    st.exception(e)
    st.stop()

st.markdown(
    "<style>.block-container{padding-top:1.2rem;padding-bottom:3rem;}"
    "div[data-testid='stMetricValue']{font-size:1.4rem;}</style>",
    unsafe_allow_html=True,
)

st.sidebar.title("📈 Nikaido Research DB")
page = st.sidebar.radio("メニュー",
                        ["✍️ トレード入力", "📋 記録一覧", "📊 分析", "🤖 AI分析", "🏷 銘柄マスタ"])


def page_input():
    st.header("✍️ トレード入力")
    c1, c2 = st.columns(2)
    with c1:
        trade_date = st.date_input("日付", value=now_jst().date())
    with c2:
        entry_time = st.time_input("エントリー時刻", value=now_jst().time())

    c3, c4 = st.columns([1, 2])
    with c3:
        code = st.text_input("銘柄コード", placeholder="例: 7746", max_chars=6)
    looked = lookup_stock(code) if code else None
    with c4:
        stock_name = st.text_input("銘柄名", value=looked or "",
                                   placeholder="未登録なら入力（次回から自動表示）")
    if code and looked:
        st.caption(f"✅ マスタ照合: {code} → {looked}")
    elif code and not looked:
        st.caption("ℹ️ 未登録コードです。名前を入力すると保存され次回から自動表示されます。")

    side = st.radio("売買区分", SIDES, horizontal=True)
    c5, c6, c7 = st.columns(3)
    with c5:
        shares = st.number_input("株数", min_value=1, value=100, step=100)
    with c6:
        in_price = st.number_input("IN価格", min_value=0.0, value=0.0, step=1.0, format="%.1f")
    with c7:
        out_price = st.number_input("OUT価格", min_value=0.0, value=0.0, step=1.0, format="%.1f")

    pnl = calc_pnl(side, shares, in_price, out_price)
    st.metric("損益（自動計算）", f"{int(pnl):+,} 円")
    st.divider()

    discovery_route = st.radio("発見経路", DISCOVERY_ROUTES)
    watchlist_timing = st.radio("監視リスト登録タイミング", WATCHLIST_TIMINGS, horizontal=True)
    entry_methods = st.multiselect("エントリー手法（複数可）", ENTRY_METHODS)

    st.write("**条件チェック**")
    conditions = []
    cols = st.columns(2)
    for i, cond in enumerate(CONDITIONS):
        with cols[i % 2]:
            if st.checkbox(cond, key=f"cond_{cond}"):
                conditions.append(cond)

    crash_personality = st.radio("急落人格分類", CRASH_PERSONALITIES)
    prior5min_state = st.radio("エントリー直前5分の状態 ⭐必須", PRIOR5MIN_STATES, index=None)
    memo = st.text_area("自由記述（反省メモ）", height=100)
    screenshot_url = st.text_input("スクショのリンク（任意）",
                                   placeholder="チャート画像のURLを貼り付け")

    if st.button("💾 保存", type="primary", use_container_width=True):
        errors = []
        if not code:
            errors.append("銘柄コード")
        if in_price <= 0 or out_price <= 0:
            errors.append("IN/OUT価格")
        if prior5min_state is None:
            errors.append("エントリー直前5分の状態（必須）")
        if errors:
            st.error("未入力: " + " / ".join(errors))
            return
        if stock_name:
            upsert_stock(code, stock_name)
        add_trade({
            "trade_date": trade_date.isoformat(), "entry_time": entry_time.strftime("%H:%M"),
            "stock_code": code, "stock_name": stock_name, "side": side, "shares": shares,
            "in_price": in_price, "out_price": out_price, "discovery_route": discovery_route,
            "watchlist_timing": watchlist_timing, "entry_methods": entry_methods,
            "conditions": conditions, "crash_personality": crash_personality,
            "prior5min_state": prior5min_state, "memo": memo, "screenshot_url": screenshot_url,
        })
        st.success(f"保存しました（損益 {int(pnl):+,}円）。続けて入力できます。")
        st.balloons()


def _fmt(df):
    return df.style.format({"平均利益": "{:,.0f}", "平均損失": "{:,.0f}", "期待値": "{:,.0f}",
                            "勝率(%)": "{:.1f}", "PF": "{:.2f}"})


def page_list():
    st.header("📋 記録一覧")
    df = get_trades_df()
    if df.empty:
        st.info("まだ記録がありません。")
        return
    show = df[["id", "trade_date", "entry_time", "stock_code", "stock_name", "side",
               "shares", "in_price", "out_price", "pnl", "discovery_route",
               "prior5min_state", "crash_personality"]].rename(columns={
        "id": "ID", "trade_date": "日付", "entry_time": "時刻", "stock_code": "コード",
        "stock_name": "銘柄", "side": "区分", "shares": "株数", "in_price": "IN",
        "out_price": "OUT", "pnl": "損益", "discovery_route": "発見経路",
        "prior5min_state": "直前5分", "crash_personality": "人格"})
    st.dataframe(show, use_container_width=True, hide_index=True)
    st.download_button("⬇️ CSVエクスポート（バックアップ）",
                       df.to_csv(index=False).encode("utf-8-sig"),
                       file_name="nikaido_trades.csv", mime="text/csv",
                       use_container_width=True)
    st.divider()
    st.subheader("詳細 / 削除")
    sel = st.selectbox("ID選択", df["id"].tolist())
    row = df[df["id"] == sel].iloc[0]
    st.write(f"**{row['stock_code']} {row['stock_name']}** / {row['side']} / 損益 {int(row['pnl']):+,}円")
    st.write("エントリー手法:", "、".join(row["entry_methods"]) or "—")
    st.write("条件:", "、".join(row["conditions"]) or "—")
    st.write("直前5分:", row["prior5min_state"], "／ 人格:", row["crash_personality"])
    if row["memo"]:
        st.write("メモ:", row["memo"])
    if row["screenshot_url"]:
        st.markdown(f"[🖼 スクショを開く]({row['screenshot_url']})")
    if st.button("🗑 このトレードを削除", type="secondary"):
        delete_trade(int(sel))
        st.warning("削除しました。再読み込みします。")
        st.rerun()


def page_analysis():
    st.header("📊 分析")
    df = get_trades_df()
    if df.empty:
        st.info("まだ記録がありません。")
        return
    min_count = st.slider("ランキングの最小サンプル数", 1, 20, 1)
    ov = overall_summary(df)
    st.subheader("総合")
    m = st.columns(3)
    m[0].metric("総トレード数", f"{int(ov['件数'])}")
    m[1].metric("勝率", f"{ov['勝率(%)']}%")
    m[2].metric("期待値", f"{int(ov['期待値']):+,}円")
    m2 = st.columns(3)
    m2[0].metric("平均利益", f"{int(ov['平均利益']):,}円")
    m2[1].metric("平均損失", f"{int(ov['平均損失']):,}円")
    pf = ov["PF"]
    m2[2].metric("PF", "∞" if pf == float("inf") else f"{pf}")
    st.divider()
    blocks = [
        ("発見経路別ランキング", lambda: group_ranking(df, "discovery_route", min_count)),
        ("エントリー手法別ランキング", lambda: list_ranking(df, "entry_methods", min_count)),
        ("条件別ランキング", lambda: list_ranking(df, "conditions", min_count)),
        ("人格分類別ランキング", lambda: group_ranking(df, "crash_personality", min_count)),
        ("直前5分の状態別ランキング", lambda: group_ranking(df, "prior5min_state", min_count)),
        ("銘柄別ランキング", lambda: group_ranking(df, "stock_name", min_count)),
        ("曜日別", lambda: group_ranking(df, "曜日", min_count)),
        ("時間帯別", lambda: group_ranking(df, "時間帯", min_count)),
    ]
    for title, fn in blocks:
        st.subheader(title)
        r = fn()
        if r is None or r.empty:
            st.caption("該当データなし（最小サンプル数を下げてください）")
        else:
            st.dataframe(_fmt(r), use_container_width=True, hide_index=True)
    st.divider()
    st.subheader("特別分析（条件別の比較）")
    for title, a, b in SPECIAL_COMPARISONS:
        st.markdown(f"**{title}**")
        st.dataframe(_fmt(special_comparison(df, a, b)), use_container_width=True, hide_index=True)


def page_ai():
    st.header("🤖 AI分析")
    st.caption("過去データから勝ち・負けパターンと直近の傾向を自動で言語化します。")
    df = get_trades_df()
    min_count = st.slider("採用する最小サンプル数", 1, 20, 3)
    if st.button("🔍 分析する", type="primary", use_container_width=True):
        for line in generate_insights(df, min_count):
            st.markdown(f"- {line}")
        st.info("※ 現在はデータ集計に基づくルールベース分析です。将来LLM呼び出しに差し替え可能な設計です。")


def page_master():
    st.header("🏷 銘柄マスタ")
    st.caption("コード→銘柄名の対応表。入力時に自動照合されます。CSV一括取込も可能。")
    with st.form("add_stock"):
        c1, c2 = st.columns([1, 2])
        code = c1.text_input("コード")
        name = c2.text_input("銘柄名")
        if st.form_submit_button("追加 / 更新"):
            if code and name:
                upsert_stock(code, name)
                st.success(f"{code} → {name} を登録しました。")
            else:
                st.error("コードと銘柄名を入力してください。")
    up = st.file_uploader("CSV一括取込（列: code, name）", type=["csv"])
    if up is not None:
        try:
            n = import_stock_csv(pd.read_csv(up, dtype=str))
            st.success(f"{n}件を取り込みました。")
        except Exception as e:
            st.error(f"取込に失敗しました: {e}")
    st.dataframe(all_stocks_df(), use_container_width=True, hide_index=True)


{
    "✍️ トレード入力": page_input,
    "📋 記録一覧": page_list,
    "📊 分析": page_analysis,
    "🤖 AI分析": page_ai,
    "🏷 銘柄マスタ": page_master,
}[page]()
