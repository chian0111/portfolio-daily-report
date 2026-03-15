#!/usr/bin/env python3
"""
📈 Portfolio Daily Report System v3
=====================================
自動抓取股價、生成 PDF 報告、AI 市場摘要，每日寄送到信箱。

功能：
  - 美股 + 台股即時報價（Yahoo Finance）
  - Private Bank 風格 PDF 報告（A4 橫向）
  - Claude AI 繁體中文市場摘要
  - 每日美股 / 台股 / 總經新聞整合
  - Gmail HTML 通知（含 PDF 附件）
  - launchd 排程自動執行（macOS）
  - 防重複執行機制

設定檔：
  ~/Desktop/config.json   — Gmail / Anthropic API Key
  ~/Desktop/portfolio.json — 持倉 / 目標金額

作者：Portfolio System
版本：v3.0
"""

import yfinance as yf
import pandas as pd
import json, os, smtplib, webbrowser, io, base64, logging
import urllib.request, urllib.error, urllib.parse
import html as html_lib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta, timezone
from zoneinfo import ZoneInfo
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.colors import HexColor
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_RIGHT, TA_LEFT
from reportlab.pdfgen import canvas
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

# ─────────────────────────────────────────────
# Timezone — always Asia/Taipei
# ─────────────────────────────────────────────
TZ = ZoneInfo("Asia/Taipei")

def now_tw():
    """Return current datetime in Asia/Taipei."""
    return datetime.now(TZ)

# ─────────────────────────────────────────────
# Logging
# ─────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)
log = logging.getLogger("portfolio")

# ─────────────────────────────────────────────
# Load config from ~/Desktop/config.json
# ─────────────────────────────────────────────
_config_path = os.path.expanduser("~/Desktop/config.json")
if os.path.exists(_config_path):
    with open(_config_path, "r") as _f:
        CONFIG = json.load(_f)
    log.info(f"Config loaded from {_config_path}")
else:
    log.warning("config.json not found — using fallback (notifications may fail)")
    CONFIG = {
        "gmail_sender":      "",
        "gmail_password":    "",
        "gmail_receiver":    "",
        "anthropic_api_key": "",
    }

# ─────────────────────────────────────────────
# Portfolio — loaded from ~/Desktop/portfolio.json
# ─────────────────────────────────────────────
_portfolio_path = os.path.expanduser("~/Desktop/portfolio.json")
if os.path.exists(_portfolio_path):
    with open(_portfolio_path, "r") as _pf:
        _pdata = json.load(_pf)
    PORTFOLIO = _pdata["holdings"]
    CASH_TWD  = _pdata.get("cash_twd", 0)
    GOAL_TWD  = _pdata.get("goal_twd", 2_000_000)
    log.info(f"Portfolio loaded from {_portfolio_path} ({len(PORTFOLIO)} holdings)")
else:
    log.error(f"portfolio.json not found at {_portfolio_path}!")
    PORTFOLIO = {}
    CASH_TWD  = 0
    GOAL_TWD  = 2_000_000


# ─────────────────────────────────────────────
# Data fetching (improved error handling)
# ─────────────────────────────────────────────
def get_usd_to_twd():
    try:
        rate = yf.Ticker("USDTWD=X").fast_info["lastPrice"]
        return round(rate, 2)
    except (KeyError, ConnectionError, ValueError) as e:
        log.warning(f"Failed to fetch USD/TWD rate: {e}, using fallback 31.88")
        return 31.88

def fetch_prices():
    symbols = list(PORTFOLIO.keys())
    prices  = {}
    try:
        tickers = yf.Tickers(" ".join(symbols))
    except Exception as e:
        log.error(f"Failed to initialize yfinance Tickers: {e}")
        return {s: {"price": None, "prev_close": None, "change_pct": 0} for s in symbols}

    for s in symbols:
        try:
            info  = tickers.tickers[s].fast_info
            price = info["lastPrice"]
            prev  = info.get("previousClose", price)
            chg   = ((price - prev) / prev * 100) if prev else 0
            prices[s] = {"price": price, "prev_close": prev, "change_pct": chg}
        except (KeyError, AttributeError, TypeError) as e:
            log.warning(f"Failed to fetch price for {s}: {e}")
            prices[s] = {"price": None, "prev_close": None, "change_pct": 0}
    return prices

def compute_pnl(prices, usd_twd):
    rows        = []
    grand_value = CASH_TWD
    grand_cost  = CASH_TWD
    for symbol, cfg in PORTFOLIO.items():
        p     = prices.get(symbol, {})
        price = p.get("price") or cfg["avg_cost"]
        rate  = usd_twd if cfg["currency"] == "USD" else 1.0
        val   = cfg["shares"] * price * rate
        cost  = cfg["shares"] * cfg["avg_cost"] * rate
        pnl   = val - cost
        pct   = (pnl / cost * 100) if cost else 0
        day_d = cfg["shares"] * price * (p.get("change_pct", 0) / 100) * rate
        grand_value += val
        grand_cost  += cost
        rows.append({
            "symbol": symbol, "shares": cfg["shares"], "price": price,
            "currency": cfg["currency"], "change_pct": p.get("change_pct", 0),
            "value_twd": val, "cost_twd": cost,
            "pnl_twd": pnl, "pnl_pct": pct, "day_delta_twd": day_d,
            "monthly_buy": cfg.get("monthly_buy", 0),
        })
    total_pnl = grand_value - grand_cost
    total_pct = (grand_value / grand_cost - 1) * 100 if grand_cost else 0
    return rows, grand_value, total_pnl, total_pct

def fetch_history_30d():
    """Fetch 30-day history for ALL stocks (US + TW)."""
    all_symbols = list(PORTFOLIO.keys())
    end   = datetime.today()
    start = end - timedelta(days=40)
    try:
        df = yf.download(all_symbols, start=start, end=end, progress=False, auto_adjust=True)["Close"]
        if isinstance(df, pd.Series):
            df = df.to_frame()
        # Drop columns with all NaN
        df = df.dropna(axis=1, how='all')
        return df
    except Exception as e:
        log.warning(f"Failed to fetch history: {e}")
        return pd.DataFrame()


# ─────────────────────────────────────────────
# 新聞抓取 — 使用 yfinance .news（最穩定）
# ─────────────────────────────────────────────
_MARKET_KEYWORDS = [
    "stock", "market", "S&P", "Nasdaq", "Dow", "rally", "sell-off",
    "earnings", "index", "shares", "equit", "bull", "bear",
    "invest", "trade", "wall street", "futures", "portfolio",
    "dividend", "IPO", "buyback", "valuation", "sector"
]
_FED_KEYWORDS = [
    "Fed", "Federal Reserve", "rate", "inflation", "CPI", "PCE", "GDP",
    "recession", "Powell", "FOMC", "tariff", "economy", "macro",
    "interest", "monetary", "fiscal", "yield", "treasury", "debt",
    "hawkish", "dovish", "quantitative", "tightening", "easing"
]
_TW_KEYWORDS = [
    "Taiwan", "TSMC", "TWD", "Taiex", "台積", "台股", "加權",
    "鴻海", "聯發科", "台灣", "外資", "央行", "證交所"
]
_NOISE_KEYWORDS = [
    # 娛樂/生活
    "horoscope", "zodiac", "celebrity", "recipe", "sport",
    "nfl", "nba", "mlb", "nhl", "soccer", "football", "basketball",
    "entertainment", "music", "movie", "film", "actor", "actress",
    "fashion", "beauty", "diet", "weight loss", "workout",
    # 低品質財經內容
    "best credit card", "best mortgage", "best savings",
    "car insurance", "home insurance", "life insurance",
    "make money fast", "get rich", "passive income tips",
    "10 stocks to buy", "5 stocks", "3 stocks",
    "here's why", "here is why", "you need to know",
    # 標題黨
    "you won't believe", "shocking", "breaking:", "just in:",
]

def _is_relevant(title, keywords):
    """標題是否含任一關鍵字（不分大小寫）。"""
    tl = title.lower()
    return any(k.lower() in tl for k in keywords)

def _not_noise(title):
    """排除明顯雜訊標題。"""
    return not _is_relevant(title, _NOISE_KEYWORDS)

def _yf_news(symbol, max_items=8):
    """用 yfinance 抓單一 ticker 的新聞，回傳 [(title, link), ...]。"""
    try:
        ticker = yf.Ticker(symbol)
        raw    = ticker.news or []
        if not raw:
            log.warning(f"[{symbol}] news is empty")
            return []
        items = []
        for n in raw:
            content = n.get("content") or {}
            title   = (content.get("title") or n.get("title") or "").strip()
            link    = (content.get("canonicalUrl",    {}).get("url")
                    or content.get("clickThroughUrl", {}).get("url")
                    or n.get("link") or n.get("url") or "#")
            if title and _not_noise(title):
                items.append((title, link))
            if len(items) >= max_items:
                break
        log.info(f"[{symbol}] got {len(items)} news items")
        return items
    except Exception as e:
        log.warning(f"yfinance news failed [{symbol}]: {e}")
        return []

def fetch_all_news(symbols):
    """
    用 yfinance 抓所有新聞，回傳：
    {
      "market":  [(title, link), ...],   # 美股大盤
      "tw":      [(title, link), ...],   # 台股
      "fed":     [(title, link), ...],   # 總經／Fed
      "stocks":  {symbol: [(title, link), ...], ...}  # 個股
    }
    每類各 5 則，過濾雜訊。
    """
    log.info("Fetching news via yfinance...")
    N = 5  # 每類顯示則數

    # ── 大盤：從 ^GSPC + ^IXIC 合併，篩市場相關標題 ──
    raw_market = _yf_news("^GSPC", 15) + _yf_news("^IXIC", 10)
    seen = set()
    market = []
    for t, l in raw_market:
        if t not in seen and (_is_relevant(t, _MARKET_KEYWORDS) or _is_relevant(t, _FED_KEYWORDS)):
            seen.add(t)
            market.append((t, l))
    # 大盤欄只留非 Fed 的
    market_only = [(t, l) for t, l in market if not _is_relevant(t, _FED_KEYWORDS)][:N]
    # 若過濾後不足，補入未分類但非雜訊的
    if len(market_only) < N:
        extra = [(t, l) for t, l in market if (t, l) not in market_only]
        market_only += extra[:N - len(market_only)]

    # ── Fed / 總經：從大盤新聞篩，不夠再從 ^TNX（10Y公債）補 ──
    fed = [(t, l) for t, l in market if _is_relevant(t, _FED_KEYWORDS)]
    if len(fed) < N:
        fed += [(t, l) for t, l in _yf_news("^TNX", 8)
                if _is_relevant(t, _FED_KEYWORDS) and t not in {x[0] for x in fed}]
    fed = fed[:N]

    # ── 台股：^TWII + 台股持倉，篩台灣相關 ──
    tw_symbols = [s for s in symbols if s.endswith(".TW") or s.endswith(".TWO")]
    raw_tw = _yf_news("^TWII", 10)
    for s in tw_symbols[:3]:
        raw_tw += _yf_news(s, 5)
    tw_seen = set()
    tw = []
    for t, l in raw_tw:
        if t not in tw_seen:
            tw_seen.add(t)
            tw.append((t, l))
    tw = tw[:N]

    # ── 個股：美股持倉各抓，篩相關標題 ──
    us_symbols = [s for s in symbols if not s.endswith(".TW") and not s.endswith(".TWO")]
    stocks = {}
    for sym in us_symbols[:6]:
        raw = _yf_news(sym, 8)
        # 個股新聞直接用，yfinance 已針對該股篩選
        if raw:
            stocks[sym] = raw[:N]
        log.info(f"  {sym}: {len(raw)} news")

    log.info(f"News done — market:{len(market_only)} tw:{len(tw)} fed:{len(fed)} stocks:{len(stocks)}")
    return {"market": market_only, "tw": tw, "fed": fed, "stocks": stocks}


# ─────────────────────────────────────────────
# Claude API — 每日市場摘要
# ─────────────────────────────────────────────
def generate_market_summary(rows, total_twd, pnl_twd, pnl_pct, usd_twd, news=None):
    """
    呼叫 Claude API 生成繁體中文每日市場摘要（含真實新聞）。
    失敗時回傳 None（不影響主流程）。
    """
    api_key = CONFIG.get("anthropic_api_key", "")
    if not api_key:
        log.warning("Anthropic API key not set — skipping market summary")
        return None

    # ── 持倉數據 ──
    holdings_lines = []
    for r in rows:
        holdings_lines.append(
            f"  • {r['symbol']}: 今日 {r['change_pct']:+.2f}%，持倉損益 TWD {r['pnl_twd']:+,.0f} ({r['pnl_pct']:+.1f}%)"
        )
    holdings_text = "\n".join(holdings_lines)

    today_str = now_tw().strftime("%Y年%m月%d日")
    progress  = min(total_twd / GOAL_TWD * 100, 100)
    remaining = GOAL_TWD - total_twd

    # ── 新聞整理成文字 ──
    news_section = ""
    if news:
        parts = []
        if news.get("market"):
            titles = "\n".join(f"    - {t}" for t, _ in news["market"])
            parts.append(f"【美股大盤新聞】\n{titles}")
        if news.get("tw"):
            titles = "\n".join(f"    - {t}" for t, _ in news["tw"])
            parts.append(f"【台股新聞】\n{titles}")
        if news.get("fed"):
            titles = "\n".join(f"    - {t}" for t, _ in news["fed"])
            parts.append(f"【總經／Fed 消息】\n{titles}")
        if news.get("stocks"):
            stock_lines = []
            for sym, items in news["stocks"].items():
                for t, _ in items:
                    stock_lines.append(f"    - [{sym}] {t}")
            if stock_lines:
                parts.append(f"【個股新聞】\n" + "\n".join(stock_lines))
        if parts:
            news_section = "\n\n" + "\n\n".join(parts)

    prompt = f"""你是一位專業的私人銀行投資顧問，請根據以下資料，用繁體中文撰寫今日的市場摘要報告。

【日期】{today_str}
【匯率】USD/TWD = {usd_twd}
【總資產】TWD {total_twd:,.0f}
【總損益】TWD {pnl_twd:+,.0f}（{pnl_pct:+.2f}%）
【目標達成】{progress:.1f}%（距目標 TWD {remaining:,.0f}）

【各持倉今日表現】
{holdings_text}{news_section}

請撰寫一份 280～350 字的市場摘要，嚴格遵守以下規則：

結構（四段，每段 2～3 句）：
1. 大盤與市場環境：今日美股、台股走勢，點出最重要的 1～2 個驅動因素
2. 個股亮點：持倉中漲跌最顯著的 2～3 檔，結合新聞說明具體原因
3. 總經觀察：Fed、利率、匯率、地緣政治等宏觀因素的潛在影響
4. 目標進度：一句精簡評述，帶出下一步值得關注的重點

寫作要求：
- 語氣專業、精煉，像高盛私人銀行週報
- 必須引用至少 2 則上方新聞的具體內容（用自己的話）
- 不要用條列式、不要重複標題文字、不要使用 emoji
- 數字要精確，引用持倉實際損益數字"""

    payload = json.dumps({
        "model":      "claude-sonnet-4-20250514",
        "max_tokens": 1500,
        "messages":   [{"role": "user", "content": prompt}]
    }).encode("utf-8")

    req = urllib.request.Request(
        "https://api.anthropic.com/v1/messages",
        data    = payload,
        headers = {
            "Content-Type":      "application/json",
            "x-api-key":         api_key,
            "anthropic-version": "2023-06-01",
        },
        method = "POST",
    )

    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            data = json.loads(resp.read().decode("utf-8"))
        summary = data["content"][0]["text"].strip()
        log.info("Market summary generated successfully")
        return summary
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="replace")
        log.error(f"Claude API HTTP error {e.code}: {body[:200]}")
        return None
    except Exception as e:
        log.error(f"Claude API error: {e}")
        return None


# ─────────────────────────────────────────────
# Charts — white Bloomberg style, font size 13
# ─────────────────────────────────────────────
def make_bar_chart(rows):
    symbols    = [r["symbol"] for r in rows]
    pnls       = [r["pnl_pct"] for r in rows]
    bar_colors = ["#007a45" if p >= 0 else "#c0392b" for p in pnls]

    fig, ax = plt.subplots(figsize=(16, 5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("#f8fafd")
    bars = ax.bar(symbols, pnls, color=bar_colors, width=0.7, zorder=3)
    ax.axhline(0, color="#c8d5e8", linewidth=1.0)
    ax.set_ylabel("Return %", color="#6b7c93", fontsize=13)
    ax.tick_params(axis='x', labelsize=13, colors="#1a2a3a")
    ax.tick_params(axis='y', labelsize=13, colors="#1a2a3a")
    for spine in ax.spines.values():
        spine.set_edgecolor("#c8d5e8")
    ax.grid(axis="y", color="#e2eaf5", zorder=0, linewidth=0.7)
    offset = max(abs(p) for p in pnls) * 0.02
    for bar, val in zip(bars, pnls):
        c = "#007a45" if val >= 0 else "#c0392b"
        ax.text(bar.get_x() + bar.get_width() / 2,
                bar.get_height() + (offset if val >= 0 else -offset * 2.5),
                f"{val:+.1f}%", ha="center",
                va="bottom" if val >= 0 else "top",
                color=c, fontsize=13, fontweight="bold")
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()
    return base64.b64encode(buf.getvalue()).decode()

def make_pie_chart(rows):
    labels  = [r["symbol"] for r in rows]
    sizes   = [r["value_twd"] for r in rows]
    palette = ["#5b9ec9","#6dbf8b","#e8915a","#9575cd","#e06c6c",
               "#4db6ac","#f0a500","#5ba4cf","#66bb6a","#ffa726","#26c6da","#ab47bc"]
    fig, ax = plt.subplots(figsize=(7, 5))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    wedges, texts, autotexts = ax.pie(
        sizes, labels=None, autopct="%1.1f%%",
        colors=palette[:len(sizes)],
        pctdistance=0.75, startangle=90,
        wedgeprops={"linewidth": 1.5, "edgecolor": "white"})
    for t in autotexts:
        t.set_color("#5c3317")
        t.set_fontsize(16)
        t.set_fontweight("bold")
        t.set_fontfamily("monospace")
    ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1.0, 0.5),
              fontsize=13, labelcolor="#1a2a3a", facecolor="white",
              edgecolor="#c8d5e8", framealpha=1,
              prop={"family": "monospace", "size": 16})
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white", transparent=False)
    plt.close()
    return base64.b64encode(buf.getvalue()).decode()

def make_history_chart(history_df):
    if history_df.empty:
        return ""
    norm    = history_df / history_df.iloc[0] * 100
    fig, ax = plt.subplots(figsize=(13, 4))
    fig.patch.set_facecolor("white")
    ax.set_facecolor("#f8fafd")
    palette = ["#0055a5","#007a45","#e67e22","#8e44ad","#c0392b",
               "#007f7f","#e74c3c","#2980b9","#27ae60","#B8924A","#1B2A4A","#6A9A5A"]
    for i, col in enumerate(norm.columns):
        ax.plot(norm.index, norm[col], linewidth=1.8,
                color=palette[i % len(palette)], label=col, alpha=0.9)
    ax.axhline(100, color="#c8d5e8", linewidth=1.0, linestyle="--")
    ax.set_ylabel("Indexed (100 = 30d ago)", color="#6b7c93", fontsize=13)
    ax.tick_params(axis='x', labelsize=13, colors="#1a2a3a")
    ax.tick_params(axis='y', labelsize=13, colors="#1a2a3a")
    for spine in ax.spines.values():
        spine.set_edgecolor("#c8d5e8")
    ax.grid(color="#e2eaf5", alpha=0.8, linewidth=0.7)
    ax.legend(fontsize=10, labelcolor="#1a2a3a", facecolor="white",
              edgecolor="#c8d5e8", ncol=4, loc="upper left", framealpha=1)
    plt.tight_layout()
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=150, bbox_inches="tight", facecolor="white")
    plt.close()
    return base64.b64encode(buf.getvalue()).decode()


# ─────────────────────────────────────────────
# HTML Dashboard — Private Bank luxury style
# ─────────────────────────────────────────────
def generate_html(rows, total_twd, pnl_twd, pnl_pct, usd_twd, bar_img, pie_img, hist_img, market_summary=None, news=None):
    progress  = min(total_twd / GOAL_TWD * 100, 100)
    remaining = GOAL_TWD - total_twd
    now       = now_tw().strftime("%Y-%m-%d %H:%M")

    pnl_color = "#4A7C59" if pnl_twd >= 0 else "#A3443D"
    ret_color = "#4A7C59" if pnl_pct >= 0 else "#A3443D"

    table_rows = ""
    for r in rows:
        chg_c = "#4A7C59" if r["change_pct"] >= 0 else "#A3443D"
        pnl_c = "#4A7C59" if r["pnl_twd"] >= 0 else "#A3443D"
        table_rows += f"""
        <tr>
          <td class="sym">{r['symbol']}</td>
          <td>{r['shares']:.4f}</td>
          <td>{r['price']:,.2f}</td>
          <td style="color:{chg_c}">{r['change_pct']:+.2f}%</td>
          <td style="color:{chg_c}">{r['day_delta_twd']:+,.0f}</td>
          <td>{r['value_twd']:,.0f}</td>
          <td>{r['cost_twd']:,.0f}</td>
          <td style="color:{pnl_c};font-weight:600">{r['pnl_twd']:+,.0f}</td>
          <td style="color:{pnl_c};font-weight:600">{r['pnl_pct']:+.2f}%</td>
          <td class="muted">{r['monthly_buy'] if r['monthly_buy'] else "—"}</td>
        </tr>"""

    hist_section = f'<img src="data:image/png;base64,{hist_img}" style="width:100%;border-radius:4px">' if hist_img else ""

    # ── 新聞區塊 HTML ──
    def _news_items_html(items):
        if not items:
            return '<p class="no-news">暫無資料</p>'
        lis = ""
        for title, link in items:
            lis += f'<li><a href="{link}" target="_blank" rel="noopener">{html_lib.escape(title)}</a></li>'
        return f"<ul>{lis}</ul>"

    if news:
        market_col  = _news_items_html(news.get("market", []))
        tw_col      = _news_items_html(news.get("tw",     []))
        fed_col     = _news_items_html(news.get("fed",    []))
        stock_items = []
        for sym, items in (news.get("stocks") or {}).items():
            for t, l in items:
                stock_items.append((f"[{sym}] {t}", l))
        stock_col = _news_items_html(stock_items)
        news_section = f"""
  <div class="section-title">📰 今日市場新聞</div>
  <div class="news-grid">
    <div class="news-col">
      <div class="news-col-title">🇺🇸 美股大盤</div>
      {market_col}
    </div>
    <div class="news-col">
      <div class="news-col-title">🇹🇼 台股</div>
      {tw_col}
    </div>
    <div class="news-col">
      <div class="news-col-title">🏦 總經／Fed</div>
      {fed_col}
    </div>
    <div class="news-col">
      <div class="news-col-title">📌 個股動態</div>
      {stock_col}
    </div>
  </div>
"""
    else:
        news_section = ""

    if market_summary:
        summary_section = f"""
  <div class="section-title">📋 今日市場摘要 · by Claude AI</div>
  <div class="summary-card">
    <div class="summary-text">{market_summary}</div>
    <div class="summary-meta">由 Claude AI 根據今日持倉數據與市場新聞自動生成 · 僅供參考，非投資建議</div>
  </div>
"""
    else:
        summary_section = ""

    return f"""<!DOCTYPE html>
<html lang="zh-TW">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Portfolio Daily Report</title>
<link href="https://fonts.googleapis.com/css2?family=Lora:ital,wght@0,400;0,600;1,400&family=DM+Sans:wght@300;400;500;700&display=swap" rel="stylesheet">
<style>
  :root {{
    --ivory: #FAF8F5;
    --navy: #1B2A4A;
    --gold: #B8924A;
    --gold-light: #D4B06C;
    --txt: #3A3A3A;
    --txt-light: #888880;
    --warm-gray: #7A7A6E;
    --line: #DDD8D0;
    --green: #4A7C59;
    --red: #A3443D;
    --zebra: #F6F4F0;
    --card: #FFFFFF;
  }}
  * {{ box-sizing:border-box; margin:0; padding:0 }}
  html, body {{
    background:var(--ivory); color:var(--txt);
    font-family:'DM Sans',sans-serif;
    font-size:15px;
    line-height:1.5;
  }}

  /* ── Header ── */
  .header {{
    border-top: 3px solid var(--gold);
    padding: 24px 40px 20px;
    display: flex; justify-content: space-between; align-items: flex-end;
    border-bottom: 1px solid var(--gold);
    max-width: 1400px; margin: 0 auto;
  }}
  .header h1 {{
    font-family: 'Times New Roman', 'Lora', serif;
    font-size: 1.5rem; font-weight: 700; color: var(--navy);
    letter-spacing: 0.04em;
  }}
  .header .sub {{ color: var(--warm-gray); font-size: 0.85rem; margin-top: 4px }}
  .header .right {{ text-align: right; color: var(--txt-light); font-size: 0.85rem }}

  .main {{ max-width: 1400px; margin: 0 auto; padding: 24px 40px }}

  /* ── Hero Metrics ── */
  .hero {{
    display: flex; justify-content: space-between; align-items: flex-end;
    padding: 20px 0 16px; border-bottom: 1px solid var(--line);
  }}
  .hero-left .label {{ font-size: 0.75rem; color: var(--txt-light); text-transform: uppercase; letter-spacing: 0.06em }}
  .hero-left .value {{ font-size: 2.2rem; font-weight: 700; color: var(--navy); font-family: 'DM Sans', sans-serif }}
  .hero-right {{ text-align: right }}
  .hero-right .label {{ font-size: 0.75rem; color: var(--txt-light); text-transform: uppercase; letter-spacing: 0.06em }}
  .hero-right .value {{ font-size: 2.2rem; font-weight: 700; font-family: 'DM Sans', sans-serif }}

  /* ── Secondary KPIs ── */
  .kpi-row {{
    display: grid; grid-template-columns: repeat(3, 1fr);
    gap: 0; padding: 18px 0; border-bottom: 1px solid var(--line);
  }}
  .kpi {{ padding: 0 8px }}
  .kpi:not(:first-child) {{ border-left: 1px solid var(--line) }}
  .kpi .label {{ font-size: 0.7rem; color: var(--txt-light); text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 4px }}
  .kpi .value {{ font-size: 1.3rem; font-weight: 700 }}

  /* ── Progress Bar ── */
  .progress-section {{ padding: 20px 0; border-bottom: 1px solid var(--line) }}
  .progress-header {{
    display: flex; justify-content: space-between; margin-bottom: 10px;
    font-size: 0.8rem; font-weight: 600;
  }}
  .progress-header .left {{ color: var(--navy) }}
  .progress-header .right {{ color: var(--txt) }}
  .progress-track {{
    background: #EDE9E3; border-radius: 6px; height: 12px; overflow: hidden;
  }}
  .progress-fill {{
    height: 100%; border-radius: 6px;
    background: linear-gradient(90deg, #1B2A4A 0%, #2E4A6A 25%, #6A6A50 50%, #A08848 75%, #D4B06C 100%);
    width: {progress:.1f}%;
    position: relative;
  }}
  .progress-fill::after {{
    content: "{progress:.1f}%";
    position: absolute; left: 12px; top: 50%; transform: translateY(-50%);
    color: white; font-size: 0.7rem; font-weight: 700;
  }}
  .progress-scale {{
    display: flex; justify-content: space-between;
    font-size: 0.65rem; color: var(--txt-light); margin-top: 4px; padding: 0 2px;
  }}

  /* ── Section Title ── */
  .section-title {{
    font-size: 0.8rem; font-weight: 700; color: var(--navy);
    text-transform: uppercase; letter-spacing: 0.05em;
    margin: 24px 0 12px; padding-bottom: 8px;
    border-bottom: 1px solid var(--line);
  }}

  /* ── Charts ── */
  .charts-grid {{ display: grid; grid-template-columns: 1.4fr 1fr; gap: 24px; margin-bottom: 8px }}
  .chart-card {{
    background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 20px;
  }}
  .chart-card img {{ width: 100%; border-radius: 4px }}
  .hist-card {{
    background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 20px; margin-bottom: 8px;
  }}

  /* ── Table ── */
  .table-card {{
    background: var(--card); border: 1px solid var(--line); border-radius: 6px; padding: 20px;
  }}
  .table-wrap {{ overflow-x: auto }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.85rem }}
  thead tr {{ border-bottom: 2px solid var(--navy) }}
  th {{
    color: var(--warm-gray); font-size: 0.7rem; font-weight: 500;
    text-transform: uppercase; letter-spacing: 0.04em;
    padding: 10px 12px; text-align: right; white-space: nowrap;
  }}
  th:first-child {{ text-align: left }}
  td {{
    padding: 10px 12px; text-align: right;
    border-bottom: 1px solid #F0ECE6; font-variant-numeric: tabular-nums;
  }}
  td:first-child {{ text-align: left }}
  tbody tr:nth-child(even) {{ background: var(--zebra) }}
  tbody tr:hover {{ background: #EFEBE4 }}
  td.sym {{ color: var(--navy); font-weight: 700 }}
  td.muted {{ color: var(--txt-light) }}

  /* ── Footer ── */
  .footer {{
    text-align: center; color: var(--txt-light); font-size: 0.75rem;
    padding: 24px 40px; border-top: 1px solid var(--line);
    max-width: 1400px; margin: 16px auto 0;
  }}

  /* ── Market Summary ── */
  .summary-card {{
    background: linear-gradient(135deg, #1B2A4A 0%, #2E4A6A 100%);
    border-radius: 8px; padding: 24px 28px; margin-bottom: 8px;
    border-left: 4px solid var(--gold);
  }}
  .summary-text {{
    color: #F0EDE8; font-size: 0.95rem; line-height: 1.85;
    font-family: 'Lora', serif; white-space: pre-wrap;
  }}
  .summary-meta {{
    color: #7A8FA8; font-size: 0.7rem; margin-top: 14px;
    padding-top: 10px; border-top: 1px solid rgba(255,255,255,0.12);
  }}

  /* ── News Grid ── */
  .news-grid {{
    display: grid; grid-template-columns: repeat(4, 1fr);
    gap: 16px; margin-bottom: 8px;
  }}
  .news-col {{
    background: var(--card); border: 1px solid var(--line);
    border-radius: 6px; padding: 16px;
  }}
  .news-col-title {{
    font-size: 0.72rem; font-weight: 700; color: var(--navy);
    text-transform: uppercase; letter-spacing: 0.04em;
    margin-bottom: 10px; padding-bottom: 6px;
    border-bottom: 1px solid var(--line);
  }}
  .news-col ul {{ list-style: none; padding: 0; margin: 0 }}
  .news-col li {{
    padding: 5px 0; border-bottom: 1px solid #F0ECE6;
    font-size: 0.78rem; line-height: 1.4;
  }}
  .news-col li:last-child {{ border-bottom: none }}
  .news-col a {{
    color: var(--txt); text-decoration: none;
    display: block;
  }}
  .news-col a:hover {{ color: var(--navy); text-decoration: underline }}
  .no-news {{ color: var(--txt-light); font-size: 0.78rem; font-style: italic }}
  @media (max-width: 900px) {{
    .news-grid {{ grid-template-columns: repeat(2, 1fr) }}
  }}
</style>
</head>
<body>

<div class="header">
  <div>
    <h1>PORTFOLIO DAILY REPORT</h1>
    <div class="sub">USD/TWD {usd_twd} &nbsp;|&nbsp; Goal TWD {GOAL_TWD:,}</div>
  </div>
  <div class="right">{now}</div>
</div>

<div class="main">

  <div class="hero">
    <div class="hero-left">
      <div class="label">Total Assets</div>
      <div class="value">TWD {total_twd:,.0f}</div>
    </div>
    <div class="hero-right">
      <div class="label">Return</div>
      <div class="value" style="color:{ret_color}">{pnl_pct:+.2f}%</div>
    </div>
  </div>

  <div class="kpi-row">
    <div class="kpi">
      <div class="label">Total P&amp;L</div>
      <div class="value" style="color:{pnl_color}">TWD {pnl_twd:+,.0f}</div>
    </div>
    <div class="kpi">
      <div class="label">Goal Progress</div>
      <div class="value" style="color:var(--navy)">{progress:.1f}%</div>
    </div>
    <div class="kpi">
      <div class="label">To Goal</div>
      <div class="value" style="color:var(--navy)">TWD {remaining:,.0f}</div>
    </div>
  </div>

  <div class="progress-section">
    <div class="progress-header">
      <span class="left">GOAL PROGRESS</span>
      <span class="right">TWD {remaining:,.0f} remaining</span>
    </div>
    <div class="progress-track"><div class="progress-fill"></div></div>
    <div class="progress-scale">
      <span>0</span><span>500K</span><span>1,000K</span><span>1,500K</span><span>2,000K</span>
    </div>
  </div>

  <div class="section-title">Return % by Symbol</div>
  <div class="charts-grid">
    <div class="chart-card">
      <img src="data:image/png;base64,{bar_img}">
    </div>
    <div class="chart-card">
      <img src="data:image/png;base64,{pie_img}">
    </div>
  </div>

  <div class="section-title">30-Day Price Trend (All Stocks, Indexed)</div>
  <div class="hist-card">
    {hist_section}
  </div>

  {news_section}

  {summary_section}

  <div class="section-title">Holdings Detail</div>
  <div class="table-card">
    <div class="table-wrap">
      <table>
        <thead><tr>
          <th style="text-align:left">Symbol</th><th>Shares</th><th>Price</th><th>Day %</th><th>Day &Delta;</th>
          <th>Value TWD</th><th>Cost TWD</th><th>P&amp;L TWD</th><th>Return %</th><th>Monthly</th>
        </tr></thead>
        <tbody>{table_rows}</tbody>
      </table>
    </div>
  </div>

</div>
<div class="footer">Data source: Yahoo Finance &nbsp;|&nbsp; For reference only, not investment advice</div>
</body></html>"""


# ─────────────────────────────────────────────
# PDF Report — Private Bank luxury style (v6)
# Uses Canvas for precise layout control
# ─────────────────────────────────────────────
def generate_pdf(rows, total_twd, pnl_twd, pnl_pct, usd_twd, bar_img_b64, pie_img_b64):
    from reportlab.lib.pagesizes import landscape
    W, H_PG = landscape(A4)   # 297 × 210 mm
    date_str  = now_tw().strftime("%Y-%m-%d")
    time_str  = now_tw().strftime("%H:%M")
    pdf_path  = os.path.expanduser(f"~/Desktop/portfolio_report_{date_str}.pdf")
    progress  = min(total_twd / GOAL_TWD * 100, 100)
    remaining = GOAL_TWD - total_twd

    total_val  = sum(r["value_twd"] for r in rows)
    total_cost = sum(r["cost_twd"] for r in rows)
    total_pl   = sum(r["pnl_twd"] for r in rows)

    def fn(v, plus=False):
        s = f"{v:,.0f}"
        return ("+" + s) if plus and v > 0 else ("-" + f"{abs(v):,.0f}") if v < 0 else s
    def fp(v):
        s = f"{v:.2f}%"
        return ("+" + s) if v > 0 else s
    def lerp_hex(a, b, t):
        ra,ga,ba = int(a[1:3],16),int(a[3:5],16),int(a[5:7],16)
        rb,gb,bb = int(b[1:3],16),int(b[3:5],16),int(b[5:7],16)
        return HexColor(f"#{int(ra+(rb-ra)*t):02x}{int(ga+(gb-ga)*t):02x}{int(ba+(bb-ba)*t):02x}")
    def rrect(c, x, y, w, h, r, fill=None, stroke=None, sw=0.3):
        p = c.beginPath(); p.roundRect(x, y, w, h, r)
        if fill: c.setFillColor(fill)
        if stroke: c.setStrokeColor(stroke); c.setLineWidth(sw)
        c.drawPath(p, fill=1 if fill else 0, stroke=1 if stroke else 0)
    def hline(c, x1, y, x2, col, w=0.3):
        c.setStrokeColor(col); c.setLineWidth(w); c.line(x1, y, x2, y)
    def vline(c, x, y1, y2, col, w=0.3):
        c.setStrokeColor(col); c.setLineWidth(w); c.line(x, y1, x, y2)

    ivory = HexColor("#FAF8F5"); navy  = HexColor("#1B2A4A")
    gold  = HexColor("#B8924A"); wgray = HexColor("#7A7A6E")
    txt   = HexColor("#3A3A3A"); txt_l = HexColor("#888880")
    ln    = HexColor("#DDD8D0"); grn   = HexColor("#4A7C59")
    rd    = HexColor("#A3443D"); trk   = HexColor("#EDE9E3")
    zebra = HexColor("#F6F4F0")

    cv = canvas.Canvas(pdf_path, pagesize=landscape(A4))
    LM = 12*mm; RM = W - 12*mm; CW = RM - LM

    # ── 固定 y 座標（從底部算起）──
    Y_HDR  = H_PG - 12*mm
    Y_KPI  = H_PG - 33*mm
    Y_PROG = H_PG - 60*mm
    Y_BODY = H_PG - 76*mm
    Y_PIE  = 44*mm           # 往上移
    Y_FOOT = 10*mm

    # 背景
    cv.setFillColor(ivory); cv.rect(0,0,W,H_PG,fill=1,stroke=0)
    cv.setFillColor(gold);  cv.rect(0,H_PG-1.2*mm,W,1.2*mm,fill=1,stroke=0)

    # ── HEADER ──
    cv.setFillColor(navy); cv.setFont("Times-Bold", 17)
    cv.drawString(LM, Y_HDR, "PORTFOLIO DAILY REPORT")
    cv.setFillColor(wgray); cv.setFont("Helvetica", 8.5)
    cv.drawRightString(RM, Y_HDR+0.5*mm, f"{date_str}   {time_str}")
    cv.setFont("Helvetica", 8.5)
    cv.drawString(LM, Y_HDR-5*mm, f"USD/TWD {usd_twd}   |   Goal TWD {GOAL_TWD:,}")
    hline(cv, LM, Y_HDR-8.5*mm, RM, gold, 0.5)

    # ── KPI ──
    cv.setFillColor(txt_l); cv.setFont("Helvetica", 8)
    cv.drawString(LM, Y_KPI, "TOTAL ASSETS")
    cv.drawRightString(RM, Y_KPI, "RETURN")

    cv.setFillColor(navy); cv.setFont("Helvetica-Bold", 24)
    cv.drawString(LM, Y_KPI-7*mm, f"TWD {total_twd:,.0f}")
    ret_col = grn if pnl_pct >= 0 else rd
    cv.setFillColor(ret_col); cv.setFont("Helvetica-Bold", 24)
    cv.drawRightString(RM, Y_KPI-7*mm, fp(pnl_pct))

    hline(cv, LM, Y_KPI-10*mm, RM, ln, 0.25)
    sec3 = [
        ("TOTAL P&L",     f"TWD {fn(pnl_twd,True)}", grn if pnl_twd>=0 else rd),
        ("GOAL PROGRESS", f"{progress:.1f}%",          navy),
        ("TO GOAL",       f"TWD {remaining:,.0f}",      navy),
    ]
    sw3 = CW/3
    for i,(lab,val,col) in enumerate(sec3):
        sx = LM + i*sw3
        if i>0: vline(cv, sx-2*mm, Y_KPI-18*mm, Y_KPI-11*mm, ln, 0.2)
        cv.setFillColor(txt_l); cv.setFont("Helvetica", 8)
        cv.drawString(sx, Y_KPI-12*mm, lab)
        cv.setFillColor(col); cv.setFont("Helvetica-Bold", 12)
        cv.drawString(sx, Y_KPI-18*mm, val)

    # ── PROGRESS BAR ──
    cv.setFillColor(navy); cv.setFont("Helvetica-Bold",8.5)
    cv.drawString(LM, Y_PROG, "GOAL PROGRESS")
    cv.setFillColor(txt_l); cv.setFont("Helvetica",8.5)
    cv.drawRightString(RM, Y_PROG, f"TWD {remaining:,.0f} remaining")

    by = Y_PROG-5*mm; bh=3.5*mm; br=1.8*mm
    fw = CW*(progress/100)
    rrect(cv, LM, by, CW, bh, br, fill=trk)
    cv.saveState()
    clip=cv.beginPath(); clip.roundRect(LM,by,fw,bh,br); cv.clipPath(clip,stroke=0)
    stops=[("#1B2A4A",0.0),("#2E4A6A",0.25),("#6A6A50",0.5),("#A08848",0.75),("#D4B06C",1.0)]
    ns=200; ssw=fw/max(ns,1)
    for step in range(ns):
        t=step/max(ns-1,1)
        for si in range(len(stops)-1):
            if t<=stops[si+1][1] or si==len(stops)-2:
                st=max(0,min(1,(t-stops[si][1])/(stops[si+1][1]-stops[si][1])))
                c2=lerp_hex(stops[si][0],stops[si+1][0],st); break
        cv.setFillColor(c2); cv.rect(LM+step*ssw,by,ssw+0.3,bh,fill=1,stroke=0)
    cv.restoreState()
    cv.setFillColor(HexColor("#FFFFFF")); cv.setFont("Helvetica-Bold",8)
    cv.drawString(LM+3*mm, by+0.8*mm, f"{progress:.1f}%")
    mk_y=by-3.5*mm
    cv.setFillColor(txt_l); cv.setFont("Helvetica",8)
    for lbl,pos in [("0",0),("500K",0.25),("1,000K",0.5),("1,500K",0.75),("2,000K",1.0)]:
        xp=LM+CW*pos
        if pos==0: cv.drawString(xp,mk_y,lbl)
        elif pos==1.0: cv.drawRightString(xp,mk_y,lbl)
        else: cv.drawCentredString(xp,mk_y,lbl); vline(cv,xp,by-0.5*mm,by-2.5*mm,ln,0.2)

    # ── 左右分欄：長條圖(左) + 表格(右) ──
    # 可用高度 = Y_BODY 到 Y_PIE+5mm
    body_h  = Y_BODY - (Y_PIE + 8*mm)   # 動態高度
    half_w  = CW/2 - 5*mm
    n_rows  = len(rows)

    # 左欄：Return by Symbol
    lx = LM
    cv.setFillColor(navy); cv.setFont("Helvetica-Bold",8.5)
    cv.drawString(lx, Y_BODY, "RETURN BY SYMBOL")
    hline(cv,lx, Y_BODY-2.5*mm, lx+half_w, ln, 0.3)

    rh = min(body_h / max(n_rows,1), 6.5*mm)
    label_w  = 18*mm
    bar_s    = lx + label_w
    bar_maxw = half_w - label_w - 20*mm

    for i,r in enumerate(rows):
        ry = Y_BODY - 4.5*mm - i*rh
        cv.setFillColor(txt); cv.setFont("Helvetica",8.5)
        cv.drawRightString(bar_s-2*mm, ry-1*mm, r["symbol"])
        capped = min(abs(r["pnl_pct"]),100)
        bw = max((capped/100)*bar_maxw, 1.5*mm)
        bc = grn if r["pnl_pct"]>=0 else rd
        rrect(cv,bar_s+0.3*mm,ry-3*mm,bw,2.8*mm,0.7*mm,fill=HexColor("#E5E0D8"))
        rrect(cv,bar_s,        ry-2.8*mm,bw,2.8*mm,0.7*mm,fill=bc)
        cv.setFillColor(bc); cv.setFont("Helvetica-Bold",8.5)
        cv.drawString(bar_s+bw+1.5*mm, ry-1*mm, fp(r["pnl_pct"]))
        if i<n_rows-1:
            hline(cv,lx,ry-rh+1.5*mm,lx+half_w,HexColor("#F0ECE6"),0.1)

    # 右欄：Holdings Table
    rx = LM + half_w + 10*mm
    cv.setFillColor(navy); cv.setFont("Helvetica-Bold",8.5)
    cv.drawString(rx, Y_BODY, "HOLDINGS DETAIL")
    hline(cv, rx, Y_BODY-2.5*mm, RM, navy, 0.5)

    col_w = [17,13,17,13,21,21,21,17]  # mm — 縮小欄寬
    cx = [rx]
    for cw in col_w[:-1]: cx.append(cx[-1]+cw*mm)
    hd = ["Symbol","Shares","Price","Day%","Val TWD","Cost TWD","P&L TWD","Ret%"]

    cv.setFillColor(wgray); cv.setFont("Helvetica",8)
    for ci,h in enumerate(hd): cv.drawString(cx[ci], Y_BODY-6*mm, h)
    hline(cv, rx, Y_BODY-8*mm, RM, ln, 0.25)

    trh = min(body_h / max(n_rows+1.5, 1), 6*mm)
    ty  = Y_BODY - 11*mm

    for i,r in enumerate(rows):
        ry2 = ty - i*trh
        if i%2==0:
            cv.setFillColor(zebra)
            cv.rect(rx-1*mm,ry2-1.5*mm,(RM-rx)+2*mm,trh,fill=1,stroke=0)
        cv.setFont("Helvetica-Bold",8.5); cv.setFillColor(navy)
        cv.drawString(cx[0],ry2,r["symbol"])
        cv.setFont("Helvetica",8.5); cv.setFillColor(txt)
        sh=r["shares"]
        cv.drawString(cx[1],ry2,f"{sh:.2f}" if sh<100 else f"{sh:.0f}")
        cv.drawString(cx[2],ry2,f"{r['price']:,.1f}")
        dc=grn if r["change_pct"]>=0 else rd
        cv.setFillColor(dc); cv.drawString(cx[3],ry2,fp(r["change_pct"]))
        cv.setFillColor(txt)
        cv.drawString(cx[4],ry2,f"{r['value_twd']:,.0f}")
        cv.drawString(cx[5],ry2,f"{r['cost_twd']:,.0f}")
        pc=grn if r["pnl_twd"]>=0 else rd
        cv.setFillColor(pc); cv.setFont("Helvetica-Bold",8.5)
        cv.drawString(cx[6],ry2,fn(r["pnl_twd"],True))
        rc=grn if r["pnl_pct"]>=0 else rd
        cv.setFillColor(rc); cv.drawString(cx[7],ry2,fp(r["pnl_pct"]))

    # TOTAL row
    sum_y = ty - n_rows*trh - 1*mm
    hline(cv,rx,sum_y+trh-1*mm,RM,navy,0.5)
    cv.setFillColor(HexColor("#F0EDE7"))
    cv.rect(rx-1*mm,sum_y-1.5*mm,(RM-rx)+2*mm,trh,fill=1,stroke=0)
    cv.setFont("Helvetica-Bold",8.5); cv.setFillColor(navy)
    cv.drawString(cx[0],sum_y,"TOTAL")
    cv.drawString(cx[4],sum_y,f"{total_val:,.0f}")
    cv.drawString(cx[5],sum_y,f"{total_cost:,.0f}")
    pc=grn if total_pl>=0 else rd; cv.setFillColor(pc)
    cv.drawString(cx[6],sum_y,fn(total_pl,True))
    t_ret=(total_pl/total_cost*100) if total_cost else 0
    cv.drawString(cx[7],sum_y,fp(t_ret))

    # ── ASSET ALLOCATION（固定在 Y_PIE）──
    hline(cv, LM, Y_PIE+2*mm, RM, ln, 0.3)
    cv.setFillColor(navy); cv.setFont("Helvetica-Bold",9)
    cv.drawString(LM, Y_PIE, "ASSET ALLOCATION")

    pie_colors=[HexColor("#1B2A4A"),HexColor("#2E5A7A"),HexColor("#4A8A6A"),
                HexColor("#6A9A5A"),HexColor("#B8924A"),HexColor("#D4B06C"),
                HexColor("#8B7355"),HexColor("#6B8E8A"),HexColor("#A0A090"),
                HexColor("#C4B8A0"),HexColor("#7A6A5A"),HexColor("#5A7A6A")]
    sdata  = sorted(enumerate(rows),key=lambda x:x[1]["value_twd"],reverse=True)
    scols  = [pie_colors[oi%len(pie_colors)] for oi,_ in sdata]
    tot_v  = sum(r["value_twd"] for _,r in sdata)

    pie_r  = 12*mm
    pie_cx = LM + 15*mm
    pie_cy = Y_PIE - 18*mm   # 往下移，讓標題和圓餅有空隙
    start  = 90
    for i,(_,r) in enumerate(sdata):
        sweep=(r["value_twd"]/tot_v)*360
        cv.setFillColor(scols[i]); cv.setStrokeColor(ivory); cv.setLineWidth(0.5)
        p=cv.beginPath(); p.moveTo(pie_cx,pie_cy)
        p.arcTo(pie_cx-pie_r,pie_cy-pie_r,pie_cx+pie_r,pie_cy+pie_r,start,sweep)
        p.close(); cv.drawPath(p,fill=1,stroke=1); start+=sweep
    cv.setFillColor(ivory); cv.circle(pie_cx,pie_cy,6.5*mm,fill=1,stroke=0)
    cv.setFillColor(navy); cv.setFont("Helvetica-Bold",9)
    cv.drawCentredString(pie_cx,pie_cy+1*mm,str(len(rows)))
    cv.setFillColor(txt_l); cv.setFont("Helvetica",7.5)
    cv.drawCentredString(pie_cx,pie_cy-3.5*mm,"holdings")

    # 圖例：2行 x 6欄，緊貼圓餅右側
    leg_x   = LM + 32*mm
    leg_y   = Y_PIE - 8*mm   # 往下，跟圓餅頂部對齊
    n       = len(sdata)
    n_cols  = 6
    col_gap = (RM - leg_x) / n_cols
    for i,(oi,r) in enumerate(sdata):
        ci = i % n_cols
        ri = i // n_cols
        lx2 = leg_x + ci*col_gap
        ly2 = leg_y - ri*11*mm
        pct = (r["value_twd"]/tot_v)*100
        cv.setFillColor(scols[i])
        cv.rect(lx2, ly2+1*mm, 3*mm, 3*mm, fill=1, stroke=0)
        cv.setFillColor(txt); cv.setFont("Helvetica-Bold",8.5)
        cv.drawString(lx2+4*mm, ly2+1.5*mm, r["symbol"])
        cv.setFillColor(txt_l); cv.setFont("Helvetica",8)
        cv.drawString(lx2+4*mm, ly2-3.5*mm, f"{pct:.1f}%")

    # ── FOOTER ──
    hline(cv,LM,Y_FOOT+2*mm,RM,ln,0.3)
    cv.setFillColor(txt_l); cv.setFont("Helvetica",9)
    cv.drawString(LM,Y_FOOT-3*mm,"Data source: Yahoo Finance  |  For reference only, not investment advice")
    cv.drawRightString(RM,Y_FOOT-3*mm,f"Generated {date_str} {time_str}")
    cv.setFillColor(gold); cv.rect(0,0,W,1.2*mm,fill=1,stroke=0)

    cv.save()
    log.info(f"PDF saved -> {pdf_path}")
    return pdf_path
    date_str  = now_tw().strftime("%Y-%m-%d")
    time_str  = now_tw().strftime("%H:%M")
    pdf_path  = os.path.expanduser(f"~/Desktop/portfolio_report_{date_str}.pdf")
    progress  = min(total_twd / GOAL_TWD * 100, 100)
    remaining = GOAL_TWD - total_twd

    total_val  = sum(r["value_twd"] for r in rows)
    total_cost = sum(r["cost_twd"] for r in rows)
    total_pl   = sum(r["pnl_twd"] for r in rows)

    # ── Helpers ──
    def fn(v, plus=False):
        s = f"{v:,.0f}"
        return ("+" + s) if plus and v > 0 else ("-" + f"{abs(v):,.0f}") if v < 0 else s

    def fp(v):
        s = f"{v:.2f}%"
        return ("+" + s) if v > 0 else s

    def lerp_hex(a, b, t):
        ra,ga,ba = int(a[1:3],16), int(a[3:5],16), int(a[5:7],16)
        rb,gb,bb = int(b[1:3],16), int(b[3:5],16), int(b[5:7],16)
        return HexColor(f"#{int(ra+(rb-ra)*t):02x}{int(ga+(gb-ga)*t):02x}{int(ba+(bb-ba)*t):02x}")

    def rrect(c, x, y, w, h, r, fill=None, stroke=None, sw=0.3):
        p = c.beginPath(); p.roundRect(x, y, w, h, r)
        if fill: c.setFillColor(fill)
        if stroke: c.setStrokeColor(stroke); c.setLineWidth(sw)
        c.drawPath(p, fill=1 if fill else 0, stroke=1 if stroke else 0)

    def hline(c, x1, y, x2, col, w=0.3):
        c.setStrokeColor(col); c.setLineWidth(w); c.line(x1, y, x2, y)

    def vline(c, x, y1, y2, col, w=0.3):
        c.setStrokeColor(col); c.setLineWidth(w); c.line(x, y1, x, y2)

    # ── Palette ──
    ivory  = HexColor("#FAF8F5")
    navy   = HexColor("#1B2A4A")
    gold   = HexColor("#B8924A")
    gold_l = HexColor("#D4B06C")
    wgray  = HexColor("#7A7A6E")
    txt    = HexColor("#3A3A3A")
    txt_l  = HexColor("#888880")
    ln     = HexColor("#DDD8D0")
    grn    = HexColor("#4A7C59")
    rd     = HexColor("#A3443D")
    trk    = HexColor("#EDE9E3")
    zebra  = HexColor("#F6F4F0")

    cv = canvas.Canvas(pdf_path, pagesize=landscape(A4))
    LM = 20*mm; RM = W - 20*mm; CW = RM - LM

    cv.setFillColor(ivory)
    cv.rect(0, 0, W, H_PG, fill=1, stroke=0)

    # Top gold rule
    cv.setFillColor(gold)
    cv.rect(0, H_PG - 1.2*mm, W, 1.2*mm, fill=1, stroke=0)

    # ── HEADER ──
    y = H_PG - 13*mm
    cv.setFillColor(navy)
    cv.setFont("Times-Bold", 20.0)
    cv.drawString(LM, y, "PORTFOLIO DAILY REPORT")

    cv.setFillColor(wgray)
    cv.setFont("Helvetica", 11.5)
    cv.drawRightString(RM, y + 0.5*mm, f"{date_str}   {time_str}")

    y -= 6.5*mm
    cv.setFillColor(wgray)
    cv.setFont("Helvetica", 11.0)
    cv.drawString(LM, y, f"USD/TWD {usd_twd}   |   Goal TWD {GOAL_TWD:,}")

    y -= 5*mm
    hline(cv, LM, y, RM, gold, 0.5)

    # ── KEY METRICS ──
    y -= 7*mm
    cv.setFillColor(txt_l)
    cv.setFont("Helvetica", 10.0)
    cv.drawString(LM, y, "TOTAL ASSETS")
    cv.drawRightString(RM, y, "RETURN")

    y -= 9*mm
    cv.setFillColor(navy)
    cv.setFont("Helvetica-Bold", 32.5)
    cv.drawString(LM, y, f"TWD {total_twd:,.0f}")

    ret_col = grn if pnl_pct >= 0 else rd
    cv.setFillColor(ret_col)
    cv.setFont("Helvetica-Bold", 32.5)
    cv.drawRightString(RM, y, fp(pnl_pct))

    y -= 11*mm
    hline(cv, LM, y + 3*mm, RM, ln, 0.25)

    sec3 = [
        ("TOTAL P&L",     f"TWD {fn(pnl_twd, True)}",  grn if pnl_twd >= 0 else rd),
        ("GOAL PROGRESS", f"{progress:.1f}%",            navy),
        ("TO GOAL",       f"TWD {remaining:,.0f}",        navy),
    ]
    sec_w = CW / 3
    for i, (lab, val, col) in enumerate(sec3):
        sx = LM + i * sec_w
        if i > 0:
            vline(cv, sx - 2*mm, y - 7*mm, y + 1*mm, ln, 0.2)
        cv.setFillColor(txt_l)
        cv.setFont("Helvetica", 10.0)
        cv.drawString(sx, y, lab)
        cv.setFillColor(col)
        cv.setFont("Helvetica-Bold", 16.0)
        cv.drawString(sx, y - 8*mm, val)

    # ── GOAL PROGRESS BAR ──
    y -= 20*mm
    cv.setFillColor(navy)
    cv.setFont("Helvetica-Bold", 10.0)
    cv.drawString(LM, y + 7*mm, "GOAL PROGRESS")
    cv.setFillColor(txt)
    cv.setFont("Helvetica", 10.0)
    cv.drawRightString(RM, y + 7*mm, f"TWD {remaining:,.0f} remaining")

    by = y; bh = 5*mm; br = 2.5*mm
    fill_w = CW * (progress / 100)
    rrect(cv, LM, by, CW, bh, br, fill=trk)

    cv.saveState()
    clip = cv.beginPath(); clip.roundRect(LM, by, fill_w, bh, br)
    cv.clipPath(clip, stroke=0)
    stops = [("#1B2A4A", 0.0), ("#2E4A6A", 0.25), ("#6A6A50", 0.5), ("#A08848", 0.75), ("#D4B06C", 1.0)]
    ns = 300; sw = fill_w / ns
    for step in range(ns):
        t = step / max(ns - 1, 1)
        for si in range(len(stops) - 1):
            if t <= stops[si+1][1] or si == len(stops) - 2:
                st = max(0, min(1, (t - stops[si][1]) / (stops[si+1][1] - stops[si][1])))
                col = lerp_hex(stops[si][0], stops[si+1][0], st); break
        cv.setFillColor(col)
        cv.rect(LM + step * sw, by, sw + 0.3, bh, fill=1, stroke=0)
    cv.restoreState()

    cv.setFillColor(HexColor("#FFFFFF"))
    cv.setFont("Helvetica-Bold", 10.0)
    cv.drawString(LM + 4*mm, by + 1.5*mm, f"{progress:.1f}%")

    y_mk = by - 5*mm
    cv.setFillColor(txt_l)
    cv.setFont("Helvetica", 10.0)
    cv.drawString(LM, y_mk, "0")
    cv.drawCentredString(LM + CW * 0.25, y_mk, "500K")
    cv.drawCentredString(LM + CW * 0.50, y_mk, "1,000K")
    cv.drawCentredString(LM + CW * 0.75, y_mk, "1,500K")
    cv.drawRightString(RM, y_mk, "2,000K")
    for pct in [0.25, 0.50, 0.75]:
        vline(cv, LM + CW * pct, by - 0.5*mm, by - 2.5*mm, ln, 0.2)

    # ── 左右分欄：左=長條圖  右=表格 ──
    y_split = y - 18*mm
    half = CW / 2 - 6*mm   # 各半寬

    # ── LEFT: RETURN BY SYMBOL ──
    lx = LM
    ly = y_split
    cv.setFillColor(navy)
    cv.setFont("Helvetica-Bold", 10.0)
    cv.drawString(lx, ly, "RETURN BY SYMBOL")
    ly -= 3*mm
    hline(cv, lx, ly, lx + half, ln, 0.3)
    ly -= 3*mm

    label_w = 22*mm
    bar_start = lx + label_w
    bar_max_w = half - label_w - 28*mm
    rh = 7*mm   # 行高放大

    for i, r in enumerate(rows):
        ry = ly - i * rh
        cv.setFillColor(txt)
        cv.setFont("Helvetica", 10.0)
        cv.drawRightString(bar_start - 3*mm, ry - 1.5*mm, r["symbol"])

        capped = min(abs(r["pnl_pct"]), 100)
        bw = max((capped / 100) * bar_max_w, 2*mm)
        bar_col = grn if r["pnl_pct"] >= 0 else rd
        rrect(cv, bar_start + 0.3*mm, ry - 4*mm, bw, 4*mm, 1*mm, fill=HexColor("#E5E0D8"))
        rrect(cv, bar_start, ry - 3.8*mm, bw, 4*mm, 1*mm, fill=bar_col)

        cv.setFillColor(bar_col)
        cv.setFont("Helvetica-Bold", 10.0)
        cv.drawString(bar_start + bw + 2*mm, ry - 1.5*mm, fp(r["pnl_pct"]))

        if i < len(rows) - 1:
            hline(cv, lx, ry - rh + 2*mm, lx + half, HexColor("#F0ECE6"), 0.1)

    # ── RIGHT: HOLDINGS TABLE ──
    rx = LM + half + 12*mm
    ry_top = y_split
    cv.setFillColor(navy)
    cv.setFont("Helvetica-Bold", 10.0)
    cv.drawString(rx, ry_top, "HOLDINGS DETAIL")
    ry_top -= 3*mm
    hline(cv, rx, ry_top, RM, navy, 0.5)

    # 右側欄寬：總寬 half，8欄
    rw = half
    # 欄位 x 位置（相對 rx）
    col_w = [20, 18, 20, 16, 24, 24, 24, 20]  # mm
    cx = [rx]
    for w in col_w[:-1]:
        cx.append(cx[-1] + w*mm)

    hd = ["Symbol", "Shares", "Price", "Day%", "Val TWD", "Cost TWD", "P&L TWD", "Ret%"]
    ry_top -= 5*mm
    cv.setFillColor(wgray)
    cv.setFont("Helvetica", 10.0)
    for ci, h in enumerate(hd):
        cv.drawString(cx[ci], ry_top, h)

    ry_top -= 2.5*mm
    hline(cv, rx, ry_top, RM, ln, 0.25)
    ry_top -= 4*mm

    trh = 6.5*mm   # 表格行高
    for i, r in enumerate(rows):
        ry2 = ry_top - i * trh
        if i % 2 == 0:
            cv.setFillColor(zebra)
            cv.rect(rx - 1*mm, ry2 - 1.5*mm, rw + 2*mm, trh, fill=1, stroke=0)

        cv.setFont("Helvetica-Bold", 10.0)
        cv.setFillColor(navy)
        cv.drawString(cx[0], ry2, r["symbol"])

        cv.setFont("Helvetica", 10.0)
        cv.setFillColor(txt)
        sh = r["shares"]
        cv.drawString(cx[1], ry2, f"{sh:.2f}" if sh < 100 else f"{sh:.0f}")
        cv.drawString(cx[2], ry2, f"{r['price']:,.1f}")

        dc = grn if r["change_pct"] >= 0 else rd
        cv.setFillColor(dc)
        cv.drawString(cx[3], ry2, fp(r["change_pct"]))

        cv.setFillColor(txt)
        cv.drawString(cx[4], ry2, f"{r['value_twd']:,.0f}")
        cv.drawString(cx[5], ry2, f"{r['cost_twd']:,.0f}")

        pc = grn if r["pnl_twd"] >= 0 else rd
        cv.setFillColor(pc)
        cv.setFont("Helvetica-Bold", 10.0)
        cv.drawString(cx[6], ry2, fn(r["pnl_twd"], True))

        rc = grn if r["pnl_pct"] >= 0 else rd
        cv.setFillColor(rc)
        cv.drawString(cx[7], ry2, fp(r["pnl_pct"]))

    # Summary row
    sum_y = ry_top - len(rows) * trh - 1*mm
    hline(cv, rx, sum_y + trh - 1*mm, RM, navy, 0.6)
    sum_y -= 1*mm
    cv.setFillColor(HexColor("#F0EDE7"))
    cv.rect(rx - 1*mm, sum_y - 1.5*mm, rw + 2*mm, trh, fill=1, stroke=0)
    cv.setFont("Helvetica-Bold", 10.0)
    cv.setFillColor(navy)
    cv.drawString(cx[0], sum_y, "TOTAL")
    cv.drawString(cx[4], sum_y, f"{total_val:,.0f}")
    cv.drawString(cx[5], sum_y, f"{total_cost:,.0f}")
    pc = grn if total_pl >= 0 else rd
    cv.setFillColor(pc)
    cv.drawString(cx[6], sum_y, fn(total_pl, True))
    t_ret = (total_pl / total_cost * 100) if total_cost else 0
    cv.drawString(cx[7], sum_y, fp(t_ret))

    # ── ALLOCATION PIE（固定在頁面底部）──
    foot_y  = 8*mm
    pie_h   = 28*mm   # 整個 pie 區塊高度
    pie_top = foot_y + 4*mm + pie_h   # 區塊頂端 y

    hline(cv, LM, pie_top + 2*mm, RM, ln, 0.3)
    cv.setFillColor(navy)
    cv.setFont("Helvetica-Bold", 10.0)
    cv.drawString(LM, pie_top, "ASSET ALLOCATION")

    pie_colors = [
        HexColor("#1B2A4A"), HexColor("#2E5A7A"), HexColor("#4A8A6A"),
        HexColor("#6A9A5A"), HexColor("#B8924A"), HexColor("#D4B06C"),
        HexColor("#8B7355"), HexColor("#6B8E8A"), HexColor("#A0A090"),
        HexColor("#C4B8A0"), HexColor("#7A6A5A"), HexColor("#5A7A6A"),
    ]
    sorted_data   = sorted(enumerate(rows), key=lambda x: x[1]["value_twd"], reverse=True)
    sorted_colors = [pie_colors[oi % len(pie_colors)] for oi, _ in sorted_data]
    total_v = sum(r["value_twd"] for _, r in sorted_data)

    # 圓餅圖：小圓，靠左
    pie_r  = 10*mm
    pie_cx = LM + 13*mm
    pie_cy = pie_top - pie_h/2

    start = 90
    for i, (_, r) in enumerate(sorted_data):
        sweep = (r["value_twd"] / total_v) * 360
        cv.setFillColor(sorted_colors[i])
        cv.setStrokeColor(ivory); cv.setLineWidth(0.5)
        p = cv.beginPath()
        p.moveTo(pie_cx, pie_cy)
        p.arcTo(pie_cx-pie_r, pie_cy-pie_r, pie_cx+pie_r, pie_cy+pie_r, start, sweep)
        p.close()
        cv.drawPath(p, fill=1, stroke=1)
        start += sweep

    cv.setFillColor(ivory)
    cv.circle(pie_cx, pie_cy, 5.5*mm, fill=1, stroke=0)
    cv.setFillColor(navy)
    cv.setFont("Helvetica-Bold", 9.0)
    cv.drawCentredString(pie_cx, pie_cy + 0.5*mm, str(len(rows)))
    cv.setFillColor(txt_l)
    cv.setFont("Helvetica", 8.0)
    cv.drawCentredString(pie_cx, pie_cy - 3.5*mm, "holdings")

    # 圖例：單行全部排開（最多 12 檔），緊貼圓餅圖右側
    leg_x   = LM + 28*mm
    leg_y   = pie_top - 5*mm
    n       = len(sorted_data)
    col_gap = (RM - leg_x) / max(n, 1)

    for i, (oi, r) in enumerate(sorted_data):
        lx2 = leg_x + i * col_gap
        pct = (r["value_twd"] / total_v) * 100

        cv.setFillColor(sorted_colors[i])
        cv.rect(lx2, leg_y + 1*mm, 3*mm, 3*mm, fill=1, stroke=0)

        cv.setFillColor(txt)
        cv.setFont("Helvetica-Bold", 9.0)
        cv.drawString(lx2 + 4*mm, leg_y + 2*mm, r["symbol"])

        cv.setFillColor(txt_l)
        cv.setFont("Helvetica", 9.0)
        cv.drawString(lx2 + 4*mm, leg_y - 3.5*mm, f"{pct:.1f}%")

    # ── FOOTER ──
    foot_y = 8*mm
    hline(cv, LM, foot_y + 2*mm, RM, ln, 0.3)
    cv.setFillColor(txt_l)
    cv.setFont("Helvetica", 10.0)
    cv.drawString(LM, foot_y - 3*mm, "Data source: Yahoo Finance  |  For reference only, not investment advice")
    cv.drawRightString(RM, foot_y - 3*mm, f"Generated {date_str} {time_str}")
    cv.setFillColor(gold)
    cv.rect(0, 0, W, 1.2*mm, fill=1, stroke=0)

    cv.save()
    log.info(f"PDF saved -> {pdf_path}")
    return pdf_path


# ─────────────────────────────────────────────
# Notifications
# ─────────────────────────────────────────────
def send_gmail(pdf_path, total_twd, pnl_twd, pnl_pct, usd_twd, market_summary=None, news=None):
    cfg    = CONFIG
    sender = cfg.get("gmail_sender", "")
    if not sender or "your_gmail" in sender:
        log.warning("Gmail not configured — skipping")
        return
    try:
        pnl_color   = "#2e7d4f" if pnl_twd >= 0 else "#c0392b"
        pnl_sign    = f"+{pnl_twd:,.0f}" if pnl_twd >= 0 else f"{pnl_twd:,.0f}"
        pnl_pct_str = f"{pnl_pct:+.2f}%"
        progress    = min(total_twd / GOAL_TWD * 100, 100)
        remaining   = GOAL_TWD - total_twd
        date_str    = now_tw().strftime("%Y-%m-%d")

        def news_rows(items, max_n=5):
            if not items:
                return "<tr><td style='color:#999;font-style:italic;padding:8px 0;font-size:17px'>暫無資料</td></tr>"
            rows = ""
            for t, l in items[:max_n]:
                import html as _h
                safe_t = _h.escape(t)
                rows += f"""<tr><td style="padding:8px 0;border-bottom:1px solid #f0ece6;font-size:17px;line-height:1.7">
                    <a href="{l}" style="color:#1B2A4A;text-decoration:none">▸ {safe_t}</a>
                </td></tr>"""
            return rows

        def stock_news_rows(stocks_dict, max_n=5):
            if not stocks_dict:
                return "<tr><td style='color:#999;font-style:italic;padding:8px 0;font-size:17px'>暫無資料</td></tr>"
            rows = ""
            count = 0
            for sym, items in stocks_dict.items():
                for t, l in items:
                    import html as _h
                    safe_t = _h.escape(t)
                    rows += f"""<tr><td style="padding:8px 0;border-bottom:1px solid #f0ece6;font-size:17px;line-height:1.7">
                        <a href="{l}" style="color:#1B2A4A;text-decoration:none">
                        <span style="color:#B8924A;font-weight:700">[{sym}]</span> {safe_t}</a>
                    </td></tr>"""
                    count += 1
                    if count >= max_n:
                        break
                if count >= max_n:
                    break
            return rows

        summary_block = ""
        if market_summary:
            import html as _h
            safe_summary = _h.escape(market_summary).replace("\n", "<br>")
            summary_block = f"""
            <tr><td style="padding:20px 0 0">
              <div style="background:linear-gradient(135deg,#1B2A4A 0%,#2E4A6A 100%);
                          border-radius:8px;padding:20px 24px;border-left:4px solid #B8924A">
                <div style="color:#B8924A;font-size:12px;font-weight:700;text-transform:uppercase;
                            letter-spacing:0.05em;margin-bottom:12px">📋 今日市場摘要 · Claude AI</div>
                <div style="color:#F0EDE8;font-size:17px;line-height:1.9">{safe_summary}</div>
                <div style="color:#7A8FA8;font-size:12px;margin-top:12px;padding-top:10px;
                            border-top:1px solid rgba(255,255,255,0.15)">
                  由 Claude AI 根據今日持倉與市場新聞自動生成 · 僅供參考，非投資建議</div>
              </div>
            </td></tr>"""

        has_news = news and any(news.get(k) for k in ["market", "tw", "fed", "stocks"])
        news_block = ""
        if has_news:
            news_block = f"""
            <tr><td style="padding:20px 0 0">
              <div style="font-size:13px;font-weight:700;color:#1B2A4A;text-transform:uppercase;
                          letter-spacing:0.05em;padding-bottom:10px;border-bottom:2px solid #1B2A4A;
                          margin-bottom:16px">📰 今日市場新聞</div>
              <!-- 第一行：美股 + 台股 -->
              <table width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px">
                <tr valign="top">
                  <td width="50%" style="padding-right:20px">
                    <div style="font-size:13px;font-weight:700;color:#B8924A;
                                margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #EDE9E3">
                      🇺🇸 美股大盤</div>
                    <table width="100%">{news_rows(news.get("market",[]))}</table>
                  </td>
                  <td width="50%">
                    <div style="font-size:13px;font-weight:700;color:#B8924A;
                                margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #EDE9E3">
                      🇹🇼 台股</div>
                    <table width="100%">{news_rows(news.get("tw",[]))}</table>
                  </td>
                </tr>
              </table>
              <!-- 第二行：Fed + 個股 -->
              <table width="100%" cellpadding="0" cellspacing="0">
                <tr valign="top">
                  <td width="50%" style="padding-right:20px">
                    <div style="font-size:13px;font-weight:700;color:#B8924A;
                                margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #EDE9E3">
                      🏦 總經／Fed</div>
                    <table width="100%">{news_rows(news.get("fed",[]))}</table>
                  </td>
                  <td width="50%">
                    <div style="font-size:13px;font-weight:700;color:#B8924A;
                                margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid #EDE9E3">
                      📌 個股動態</div>
                    <table width="100%">{stock_news_rows(news.get("stocks",{}))}</table>
                  </td>
                </tr>
              </table>
            </td></tr>"""

        html_body = f"""<!DOCTYPE html>
<html><head><meta charset="UTF-8"></head>
<body style="margin:0;padding:0;background:#FAF8F5;font-family:'Helvetica Neue',Arial,sans-serif">
<table width="100%" cellpadding="0" cellspacing="0" style="background:#FAF8F5;padding:24px 0">
<tr><td align="center">
<table width="620" cellpadding="0" cellspacing="0"
       style="background:#ffffff;border-radius:8px;overflow:hidden;
              box-shadow:0 2px 12px rgba(0,0,0,0.08)">

  <!-- Header -->
  <tr><td style="background:#1B2A4A;padding:24px 32px;border-bottom:3px solid #B8924A">
    <div style="color:#B8924A;font-size:11px;font-weight:700;letter-spacing:0.1em;
                text-transform:uppercase;margin-bottom:4px">Daily Portfolio Report</div>
    <div style="color:#ffffff;font-size:22px;font-weight:700">{date_str}</div>
  </td></tr>

  <!-- KPIs -->
  <tr><td style="padding:24px 32px;border-bottom:1px solid #EDE9E3">
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>
        <td style="text-align:center;padding:0 8px;border-right:1px solid #EDE9E3">
          <div style="font-size:10px;color:#999;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px">Total Assets</div>
          <div style="font-size:22px;font-weight:700;color:#1B2A4A">TWD {total_twd:,.0f}</div>
        </td>
        <td style="text-align:center;padding:0 8px;border-right:1px solid #EDE9E3">
          <div style="font-size:10px;color:#999;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px">Total P&amp;L</div>
          <div style="font-size:22px;font-weight:700;color:{pnl_color}">TWD {pnl_sign} ({pnl_pct_str})</div>
        </td>
        <td style="text-align:center;padding:0 8px;border-right:1px solid #EDE9E3">
          <div style="font-size:10px;color:#999;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px">USD/TWD</div>
          <div style="font-size:22px;font-weight:700;color:#1B2A4A">{usd_twd}</div>
        </td>
        <td style="text-align:center;padding:0 8px">
          <div style="font-size:10px;color:#999;text-transform:uppercase;letter-spacing:0.05em;margin-bottom:4px">Goal Progress</div>
          <div style="font-size:22px;font-weight:700;color:#1B2A4A">{progress:.1f}%</div>
          <div style="font-size:11px;color:#999">TWD {remaining:,.0f} to go</div>
        </td>
      </tr>
    </table>
    <!-- Progress bar -->
    <div style="margin-top:16px;background:#EDE9E3;border-radius:4px;height:8px;overflow:hidden">
      <div style="width:{progress:.1f}%;height:100%;
                  background:linear-gradient(90deg,#1B2A4A,#B8924A);border-radius:4px"></div>
    </div>
  </td></tr>

  <!-- News + Summary -->
  <tr><td style="padding:0 32px 24px">
    <table width="100%" cellpadding="0" cellspacing="0">
      {news_block}
      {summary_block}
    </table>
  </td></tr>

  <!-- Footer -->
  <tr><td style="background:#F6F4F0;padding:16px 32px;border-top:1px solid #EDE9E3;
                 text-align:center;font-size:11px;color:#aaa">
    Data source: Yahoo Finance &nbsp;·&nbsp; For reference only, not investment advice<br>
    <span style="color:#B8924A">See attached PDF for full report</span>
  </td></tr>

</table>
</td></tr>
</table>
</body></html>"""

        msg         = MIMEMultipart("alternative")
        msg["From"] = sender
        msg["To"]   = cfg["gmail_receiver"]
        msg["Subject"] = f"📈 Daily Portfolio Report {date_str} | TWD {total_twd:,.0f}"
        msg.attach(MIMEText(html_body, "html", "utf-8"))

        # PDF 附件
        with open(pdf_path, "rb") as f:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header("Content-Disposition",
                            f"attachment; filename={os.path.basename(pdf_path)}")
            msg.attach(part)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender, cfg["gmail_password"])
            server.send_message(msg)
        log.info("Gmail (HTML) sent successfully!")
    except smtplib.SMTPAuthenticationError:
        log.error("Gmail authentication failed — check password/app password")
    except smtplib.SMTPException as e:
        log.error(f"Gmail SMTP error: {e}")
    except Exception as e:
        log.error(f"Gmail unexpected error: {e}")


def save_log(rows, total_twd, usd_twd):
    log_file = os.path.expanduser("~/Desktop/portfolio_log.json")
    today    = now_tw().strftime("%Y-%m-%d")
    snap     = {
        "date": today,
        "usd_twd": usd_twd,
        "total_twd": round(total_twd),
        "goal_pct": round(total_twd / GOAL_TWD * 100, 2),
        "holdings": {r["symbol"]: round(r["value_twd"]) for r in rows},
    }
    log_data = []
    if os.path.exists(log_file):
        try:
            with open(log_file) as f:
                log_data = json.load(f)
        except (json.JSONDecodeError, IOError) as e:
            log.warning(f"Failed to read existing log: {e}")
    log_data = [e for e in log_data if e["date"] != today]
    log_data.append(snap)
    with open(log_file, "w") as f:
        json.dump(log_data, f, indent=2, ensure_ascii=False)
    log.info(f"Log saved -> {log_file}")


# ─────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────
def _already_ran_today() -> bool:
    """防止 launchd 在同一時段重複觸發。"""
    lock_path = os.path.expanduser("~/Desktop/.portfolio_last_run")
    today     = now_tw().strftime("%Y-%m-%d")
    curr_hour = now_tw().hour
    # 分兩個時段：白天(0–13) / 晚上(14–23)
    curr_block = "day" if curr_hour < 14 else "night"
    if os.path.exists(lock_path):
        try:
            with open(lock_path) as f:
                data = json.load(f)
            if data.get("date") == today and data.get("block") == curr_block:
                log.warning(f"Already ran today [{curr_block}], skipping.")
                return True
        except Exception:
            pass
    with open(lock_path, "w") as f:
        json.dump({"date": today, "block": curr_block}, f)
    return False


def main():
    if _already_ran_today():
        print("⏭️  Already ran in this time block today. Exiting.")
        return

    t = now_tw()
    print("=" * 55)
    print("  📈  PORTFOLIO SYSTEM v2".center(55))
    print(f"  {t.strftime('%Y-%m-%d %H:%M:%S')} (Asia/Taipei)".center(55))
    print("=" * 55)

    log.info("Fetching exchange rate...")
    usd_twd = get_usd_to_twd()
    log.info(f"USD/TWD = {usd_twd}")

    log.info("Fetching stock prices...")
    prices = fetch_prices()

    log.info("Computing P&L...")
    rows, total_twd, pnl_twd, pnl_pct = compute_pnl(prices, usd_twd)

    log.info("Fetching 30-day history (US + TW)...")
    history = fetch_history_30d()

    log.info("Generating charts...")
    bar_img  = make_bar_chart(rows)
    pie_img  = make_pie_chart(rows)
    hist_img = make_history_chart(history)

    log.info("Fetching news...")
    symbols = list(PORTFOLIO.keys())
    news = fetch_all_news(symbols)

    log.info("Generating market summary (Claude AI)...")
    market_summary = generate_market_summary(rows, total_twd, pnl_twd, pnl_pct, usd_twd, news)

    log.info("Generating HTML dashboard...")
    html      = generate_html(rows, total_twd, pnl_twd, pnl_pct, usd_twd, bar_img, pie_img, hist_img, market_summary, news)
    html_path = os.path.expanduser("~/Desktop/portfolio_dashboard.html")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(html)
    log.info(f"HTML saved -> {html_path}")
    webbrowser.open(f"file://{html_path}")

    log.info("Generating PDF report...")
    pdf_path = generate_pdf(rows, total_twd, pnl_twd, pnl_pct, usd_twd, bar_img, pie_img)

    log.info("Sending notifications...")
    send_gmail(pdf_path, total_twd, pnl_twd, pnl_pct, usd_twd, market_summary, news)

    save_log(rows, total_twd, usd_twd)

    print(f"\n{'─'*55}")
    print(f"  💵 Total Assets : TWD {total_twd:>15,.0f}")
    print(f"  📊 Total P&L    : TWD {pnl_twd:>+15,.0f} ({pnl_pct:+.2f}%)")
    print(f"  🎯 Goal         : {min(total_twd/GOAL_TWD*100,100):.1f}% — 差 TWD {GOAL_TWD-total_twd:,.0f}")
    print(f"{'─'*55}")
    print("\n✅ Done!\n")

if __name__ == "__main__":
    main()
